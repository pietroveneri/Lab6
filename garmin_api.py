"""Accesso Garmin Connect e lettura dei BPM originali di un'attività FIT.

Compatibilità: Python >= 3.12, garminconnect 0.3.17,
garmin-fit-sdk 21.218.0. Le dipendenze sono caricate solo quando servono.

Fonti:
https://github.com/cyberjunky/python-garminconnect
https://github.com/garmin/fit-python-sdk
"""

from __future__ import annotations

import io
import math
import os
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from getpass import getpass
from pathlib import Path
from typing import Any


def _garmin_classes() -> tuple[Any, Any, Any, Any]:
    if sys.version_info < (3, 12):
        raise RuntimeError("L'accesso Garmin richiede Python 3.12 o successivo.")
    try:
        from garminconnect import (
            Garmin,
            GarminConnectAuthenticationError,
            GarminConnectConnectionError,
            GarminConnectTooManyRequestsError,
        )
    except (ImportError, OSError):
        raise RuntimeError(
            "Dipendenza Garmin non disponibile: installare garminconnect==0.3.17 "
            "con il file requirements.txt."
        ) from None
    return (
        Garmin,
        GarminConnectAuthenticationError,
        GarminConnectConnectionError,
        GarminConnectTooManyRequestsError,
    )


def login_garmin(token_dir: Path, email: str | None = None) -> Any:
    """Riusa la sessione salvata oppure richiede credenziali e codice MFA.

    Le credenziali non sono accettate dalla riga di comando né stampate.
    Le sessioni 0.3.x sono persistite da login(tokenstore), senza garth.
    Gli errori di autenticazione sono tradotti senza esporre la risposta.
    """
    Garmin, AuthError, ConnectionError, RateError = _garmin_classes()
    token_dir = Path(token_dir).expanduser()
    tokenstore = str(token_dir)

    if (token_dir / "garmin_tokens.json").is_file():
        client = Garmin()
        try:
            client.login(tokenstore)
            return client
        except AuthError:
            # Sessione scaduta o revocata: esegui un nuovo login una sola volta.
            pass
        except RateError:
            raise RuntimeError(
                "Garmin ha limitato gli accessi. Attendere prima di riprovare."
            ) from None
        except ConnectionError:
            raise RuntimeError(
                "Impossibile contattare Garmin o ripristinare la sessione. "
                "Verificare la connessione e riprovare."
            ) from None
        except Exception:
            raise RuntimeError(
                "Impossibile ripristinare la sessione Garmin salvata."
            ) from None

    address = (email or input("Email Garmin Connect: ")).strip()
    if not address:
        raise ValueError("L'email Garmin Connect è necessaria per il primo accesso.")
    password = getpass("Password Garmin Connect: ")
    if not password:
        raise ValueError("La password Garmin Connect non può essere vuota.")
    client = Garmin(
        address,
        password,
        prompt_mfa=lambda: getpass("Codice MFA Garmin: ").strip(),
    )
    # Non conservare un riferimento aggiuntivo alla password nel modulo.
    password = None
    try:
        client.login(tokenstore)
    except AuthError:
        raise RuntimeError(
            "Accesso Garmin non riuscito. Verificare email, password e codice MFA."
        ) from None
    except RateError:
        raise RuntimeError(
            "Garmin ha limitato gli accessi. Attendere prima di riprovare."
        ) from None
    except ConnectionError:
        raise RuntimeError(
            "Impossibile contattare Garmin. Verificare la connessione e riprovare."
        ) from None
    except Exception:
        raise RuntimeError("Accesso Garmin non riuscito.") from None
    return client


def fetch_activity_fit(
    activity_id: str,
    token_dir: Path,
    cache_path: Path,
    email: str | None = None,
) -> bytes:
    """Scarica il FIT originale tramite l'API e lo salva in cache_path.

    ORIGINAL restituisce un archivio ZIP. L'estrazione avviene in memoria e
    richiede un solo FIT, così nessuna attività viene scelta implicitamente.
    cache_path contiene il FIT estratto, non il contenuto ZIP.
    """
    identifier = str(activity_id).strip()
    if not identifier.isascii() or not identifier.isdecimal() or int(identifier) <= 0:
        raise ValueError("L'ID dell'attività Garmin deve essere un intero positivo.")
    client = login_garmin(token_dir, email)
    Garmin, AuthError, ConnectionError, RateError = _garmin_classes()
    try:
        archive = client.download_activity(
            identifier, dl_fmt=Garmin.ActivityDownloadFormat.ORIGINAL
        )
    except AuthError:
        raise RuntimeError(
            "Garmin non ha autorizzato il download. Ripetere l'accesso."
        ) from None
    except RateError:
        raise RuntimeError(
            "Garmin ha limitato le richieste. Attendere prima di riprovare."
        ) from None
    except ConnectionError:
        raise RuntimeError(
            "Download Garmin non riuscito. Verificare l'ID dell'attività, "
            "la sua sincronizzazione in Garmin Connect e la connessione."
        ) from None
    except Exception:
        raise RuntimeError("Download dell'attività Garmin non riuscito.") from None

    try:
        with zipfile.ZipFile(io.BytesIO(archive)) as archive_file:
            candidates = [
                item
                for item in archive_file.infolist()
                if not item.is_dir() and Path(item.filename).suffix.lower() == ".fit"
            ]
            if not candidates:
                raise ValueError("L'archivio Garmin non contiene alcun file FIT.")
            if len(candidates) != 1:
                raise ValueError(
                    "L'archivio Garmin contiene più file FIT: "
                    "l'attività da utilizzare è ambigua."
                )
            data = archive_file.read(candidates[0])
    except (zipfile.BadZipFile, TypeError, EOFError, RuntimeError, NotImplementedError):
        raise ValueError("Garmin non ha restituito un archivio FIT ZIP valido.") from None

    destination = Path(cache_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=destination.parent, suffix=".fit.tmp", delete=False
        ) as temporary:
            temporary_path = Path(temporary.name)
            temporary.write(data)
        os.replace(temporary_path, destination)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return data


def read_fit_samples(data: bytes) -> list[tuple[float, int]]:
    """Restituisce coppie (epoch UTC, BPM) dei record FIT originali.

    Non calcola medie, interpolazioni o valori sostitutivi. La fusione
    automatica di hr_mesgs è disabilitata perché l'SDK calcola medie e
    riempie intervalli durante quella fusione. Il CRC resta abilitato.
    """
    try:
        from garmin_fit_sdk import Decoder, Stream
    except ImportError:
        raise RuntimeError(
            "Dipendenza FIT non disponibile: installare garmin-fit-sdk==21.218.0 "
            "con il file requirements.txt."
        ) from None

    try:
        decoder = Decoder(Stream.from_byte_array(bytearray(data)))
        messages, errors = decoder.read(enable_crc_check=True, merge_heart_rates=False)
    except Exception:
        raise ValueError("Impossibile decodificare il file FIT Garmin.") from None
    if errors:
        raise ValueError(
            "Il file FIT Garmin è danneggiato o non decodificabile "
            "(verifica del formato/CRC non superata)."
        )

    samples: list[tuple[float, int]] = []
    for record in messages.get("record_mesgs", []):
        timestamp = record.get("timestamp")
        heart_rate = record.get("heart_rate")
        if not isinstance(timestamp, datetime):
            continue
        if isinstance(heart_rate, bool) or not isinstance(heart_rate, int):
            continue
        # 255 è il valore FIT uint8 non valido; 0 non è un BPM misurato valido.
        if not 1 <= heart_rate <= 254:
            continue
        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=timezone.utc)
        else:
            timestamp = timestamp.astimezone(timezone.utc)
        epoch = timestamp.timestamp()
        if math.isfinite(epoch):
            samples.append((epoch, heart_rate))
    samples.sort(key=lambda sample: sample[0])
    if not samples:
        raise ValueError(
            "Il file FIT non contiene record con timestamp e BPM validi. "
            "Registrare una attività temporizzata con la frequenza cardiaca attiva."
        )
    return samples
