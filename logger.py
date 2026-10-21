"""Lab 6: micro:bit ogni 10 s e BPM Garmin abbinati per timestamp.

Python 3.12+. Nessuna analisi statistica. Garmin Connect viene letto DOPO
la sincronizzazione dell'attivita', tramite il client non ufficiale.
"""

from __future__ import annotations

import argparse
import asyncio
import bisect
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
import re
import sys
import time
import zipfile
from xml.sax.saxutils import escape

UART_SERVICE = "6e400001-b5a3-f393-e0a9-e50e24dcca9e"
# Stesso UUID del BLEGW.html fornito. Verifichiamo anche le proprieta' GATT.
UART_NOTIFY = "6e400002-b5a3-f393-e0a9-e50e24dcca9e"
INTERVAL = 10.0


def utc_iso(epoch: float) -> str:
    return datetime.fromtimestamp(epoch, timezone.utc).isoformat(timespec="milliseconds")


def atomic_json(path: Path, value: object) -> None:
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def numeric(text: str) -> int | float | None:
    try:
        value = float(text)
    except ValueError:
        return None
    if not math.isfinite(value):
        return None
    return int(value) if value.is_integer() else value


def parse_microbit(text: str) -> dict | None:
    """Formati del gateway: 'BPM 73' oppure 'segnale filtrato 73'."""
    fields = text.strip().split()
    signal = filtered = None
    if len(fields) == 2 and fields[0].upper() == "BPM":
        bpm = numeric(fields[1])
    elif len(fields) >= 3:
        signal, filtered, bpm = (numeric(v) for v in fields[:3])
        if signal is None or filtered is None:
            return None
    else:
        return None
    if not isinstance(bpm, int) or not 1 <= bpm <= 300:
        return None
    return {"bpm": bpm, "signal": signal, "filtered": filtered, "raw": text.strip()}


class Framer:
    """packet: gateway esistente; line: messaggi terminati da LF, anche frammentati."""

    def __init__(self, mode: str):
        self.mode = mode
        self.buffer = bytearray()

    def feed(self, data: bytes) -> list[str]:
        if self.mode == "packet":
            return data.decode("utf-8", errors="replace").splitlines()
        self.buffer.extend(data)
        if len(self.buffer) > 65536:
            self.buffer.clear()
            raise ValueError("Messaggio UART troppo lungo: manca il terminatore LF.")
        lines = []
        while b"\n" in self.buffer:
            line, _, rest = self.buffer.partition(b"\n")
            self.buffer = bytearray(rest)
            lines.append(line.decode("utf-8", errors="replace").rstrip("\r"))
        return lines


@dataclass(frozen=True)
class Sample:
    mono: float
    epoch: float
    values: dict


def nearest_microbit(samples: deque, target: float, tolerance: float) -> Sample | None:
    candidates = [s for s in samples if abs(s.mono - target) <= tolerance]
    # Parita': scegli il precedente. Nessun riuso di un campione di una riga vecchia.
    return min(candidates, key=lambda s: (abs(s.mono - target), s.mono), default=None)


def excel_column(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xml_cell(ref: str, value: object, style: int = 0) -> str:
    if value is None:
        return f'<c r="{ref}"/>'
    if isinstance(value, datetime):
        # Excel non memorizza il fuso orario: le intestazioni dichiarano UTC.
        epoch = value.astimezone(timezone.utc).timestamp()
        return f'<c r="{ref}" s="2"><v>{epoch / 86400 + 25569:.12f}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(value):
            raise ValueError("Valore non finito nel file Excel.")
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    # inlineStr conserva i messaggi originali come testo, mai come formule Excel.
    clean = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(value))
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{escape(clean)}</t></is></c>'


def write_xlsx(path: Path, session: dict) -> None:
    """File .xlsx standard senza dipendenze per Excel; nessuna formula o grafico."""
    headers = [
        "N", "Istante previsto UTC", "Microbit ricevuto UTC", "Microbit BPM",
        "Segnale", "Segnale filtrato", "Garmin originale UTC", "Garmin allineato UTC",
        "Garmin BPM", "Scarto Garmin - Microbit (s)", "Stato", "Messaggio microbit",
    ]
    grid = [headers]
    date = lambda epoch: datetime.fromtimestamp(epoch, timezone.utc) if epoch is not None else None
    for row in session["rows"]:
        grid.append([
            row["n"], date(row["target_epoch"]), date(row.get("micro_epoch")),
            row.get("micro_bpm"), row.get("signal"), row.get("filtered"),
            date(row.get("garmin_original_epoch")), date(row.get("garmin_epoch")),
            row.get("garmin_bpm"), row.get("delta_seconds"), row["status"], row.get("raw"),
        ])
    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    rows_xml = []
    for r, values in enumerate(grid, 1):
        cells = "".join(xml_cell(f"{excel_column(c)}{r}", v, 1 if r == 1 else 0)
                        for c, v in enumerate(values, 1))
        rows_xml.append(f'<row r="{r}" ht="{32 if r == 1 else 21}" customHeight="1">{cells}</row>')
    widths = [7, 26, 26, 15, 13, 18, 26, 26, 15, 23, 30, 40]
    cols = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                   for i, w in enumerate(widths, 1))
    sheet = (f'<worksheet {ns}><sheetViews><sheetView workbookViewId="0">'
             '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
             '</sheetView></sheetViews><cols>' + cols + '</cols><sheetData>' +
             "".join(rows_xml) + f'</sheetData><autoFilter ref="A1:L{len(grid)}"/></worksheet>')
    notes = [["Parametro", "Valore"]] + [[k, str(v)] for k, v in session["metadata"].items()]
    note_xml = "".join(f'<row r="{r}">' + "".join(xml_cell(f"{excel_column(c)}{r}", v, 1 if r == 1 else 0)
                       for c, v in enumerate(vals, 1)) + '</row>' for r, vals in enumerate(notes, 1))
    metadata_sheet = (f'<worksheet {ns}><cols><col min="1" max="1" width="35" customWidth="1"/>'
                      '<col min="2" max="2" width="100" customWidth="1"/></cols>'
                      f'<sheetData>{note_xml}</sheetData></worksheet>')
    styles = (f'<styleSheet {ns}><numFmts count="1"><numFmt numFmtId="164" '
              'formatCode="yyyy-mm-dd hh:mm:ss.000"/></numFmts>'
              '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
              '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>'
              '<fills count="3"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FF24566F"/>'
              '<bgColor indexed="64"/></patternFill></fill></fills>'
              '<borders count="1"><border/></borders><cellStyleXfs count="1">'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1">'
              '<alignment wrapText="1" vertical="center"/></xf>'
              '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
              '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    content_types = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                     '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                     '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    rel_ns = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    workbook = (f'<workbook {ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Misure" sheetId="1" r:id="rId1"/>'
                '<sheet name="Impostazioni" sheetId="2" r:id="rId2"/></sheets></workbook>')
    relationships = (f'<Relationships {rel_ns}><Relationship Id="rId1" Type="{office_rel}worksheet" Target="worksheets/sheet1.xml"/>'
                     f'<Relationship Id="rId2" Type="{office_rel}worksheet" Target="worksheets/sheet2.xml"/>'
                     f'<Relationship Id="rId3" Type="{office_rel}styles" Target="styles.xml"/></Relationships>')
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in {
            "[Content_Types].xml": content_types,
            "_rels/.rels": f'<Relationships {rel_ns}><Relationship Id="rId1" Type="{office_rel}officeDocument" Target="xl/workbook.xml"/></Relationships>',
            "xl/workbook.xml": workbook, "xl/_rels/workbook.xml.rels": relationships,
            "xl/styles.xml": styles, "xl/worksheets/sheet1.xml": sheet,
            "xl/worksheets/sheet2.xml": metadata_sheet,
        }.items():
            archive.writestr(name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + value)
    os.replace(tmp, path)


def checkpoint(output: Path, session_path: Path, session: dict) -> bool:
    atomic_json(session_path, session)
    try:
        write_xlsx(output, session)
    except PermissionError:
        print("Excel e' aperto: chiudilo. I dati sono salvati nel JSON; usa il comando export.", file=sys.stderr)
        return False
    return True


async def discover(args):
    try:
        from bleak import BleakScanner
    except ImportError as exc:
        raise RuntimeError("Installa le dipendenze con: python -m pip install -r requirements.txt") from exc
    devices = await BleakScanner.discover(timeout=args.scan_seconds, return_adv=True)
    if args.command == "scan":
        for device, adv in devices.values():
            print(f"{device.address}  {adv.local_name or device.name or '(senza nome)'}")
        return None
    matches = []
    for device, adv in devices.values():
        name = adv.local_name or device.name or ""
        if args.address:
            match = device.address.casefold() == args.address.casefold()
        else:
            match = args.name.casefold() in name.casefold()
        if match:
            matches.append(device)
    if len(matches) != 1:
        raise RuntimeError(f"Trovati {len(matches)} micro:bit compatibili con la selezione. Usa scan e --address per sceglierne uno.")
    return matches[0]


async def record(args) -> None:
    try:
        from bleak import BleakClient
    except ImportError as exc:
        raise RuntimeError("Installa le dipendenze con: python -m pip install -r requirements.txt") from exc

    output = args.output.resolve()
    session_path = output.with_suffix(".session.json")
    raw_path = output.with_suffix(".microbit.jsonl")
    for path in (output, session_path, raw_path):
        if path.exists():
            raise RuntimeError(f"Il file esiste gia': {path.name}. Scegli un nuovo --output.")
    output.parent.mkdir(parents=True, exist_ok=True)
    device = await discover(args)
    buffer = deque()
    framer = Framer(args.framing)
    received = asyncio.Event()
    disconnected = asyncio.Event()
    anchor_mono, anchor_epoch = time.monotonic(), time.time()
    session = {"schema_version": 1, "metadata": {
        "intervallo_secondi": INTERVAL, "microbit": device.address,
        "microbit_uuid": args.characteristic or "rilevato nel servizio UART",
        "riscaldamento_secondi": args.warmup, "tolleranza_griglia_secondi": args.tolerance,
        "timestamp_microbit": "Ricezione PC UTC; il firmware del gateway non invia un timestamp della misura",
        "timestamp_garmin": "Originale nel file FIT dell'attivita'; da scaricare dopo la sincronizzazione",
        "sincronizzazione": "Abbinamento temporale entro tolleranza; simultaneita' fisica non certificata",
        "origine_garmin": "Non ancora scaricata", "analisi_statistiche": "Nessuna",
        "framing": args.framing, "stato_registrazione": "in corso",
    }, "rows": []}
    checkpoint(output, session_path, session)

    with raw_path.open("x", encoding="utf-8") as raw_file:
        def notification(_characteristic, payload: bytearray):
            mono = time.monotonic()
            epoch = anchor_epoch + mono - anchor_mono
            raw_file.write(json.dumps({"utc": utc_iso(epoch), "epoch": epoch,
                                       "payload_hex": bytes(payload).hex(),
                                       "text": bytes(payload).decode("utf-8", errors="replace")}, ensure_ascii=False) + "\n")
            raw_file.flush()
            try:
                lines = framer.feed(bytes(payload))
            except ValueError as exc:
                print(str(exc), file=sys.stderr)
                return
            for text in lines:
                values = parse_microbit(text)
                if values and mono >= ready_at:
                    buffer.append(Sample(mono, epoch, values))
                    received.set()

        try:
            async with BleakClient(device, pair=args.pair, timeout=60,
                                   disconnected_callback=lambda _: disconnected.set()) as client:
                service = client.services.get_service(UART_SERVICE)
                if service is None:
                    raise RuntimeError("Servizio UART assente. Attiva Bluetooth UART nel firmware micro:bit.")
                chars = [c for c in service.characteristics if {"notify", "indicate"} & set(c.properties)]
                if args.characteristic:
                    chars = [c for c in chars if c.uuid.casefold() == args.characteristic.casefold()]
                preferred = [c for c in chars if c.uuid.casefold() == UART_NOTIFY]
                if preferred:
                    chars = preferred
                if len(chars) != 1:
                    raise RuntimeError("Caratteristica UART notificabile ambigua/assente. Specifica --characteristic.")
                session["metadata"]["microbit_uuid"] = chars[0].uuid
                ready_at = time.monotonic() + args.warmup
                await client.start_notify(chars[0], notification)
                print(f"Connesso. Riscaldamento {args.warmup:g} s, poi attesa del primo BPM valido.")
                deadline = ready_at + 60
                while not received.is_set():
                    if disconnected.is_set():
                        raise RuntimeError("Micro:bit disconnesso prima dell'inizio della prova.")
                    if time.monotonic() > deadline:
                        raise RuntimeError("Nessun BPM valido. Controlla firmware/formato UART e --framing.")
                    await asyncio.sleep(0.1)
                start = buffer[0]
                session["metadata"]["inizio_utc"] = utc_iso(start.epoch)
                for index in range(args.samples):
                    target = start.mono + index * INTERVAL
                    target_epoch = start.epoch + index * INTERVAL
                    # Aspetta anche eventuali campioni appena successivi all'istante.
                    await asyncio.sleep(max(0, target + args.tolerance - time.monotonic()))
                    if disconnected.is_set():
                        raise RuntimeError("Micro:bit disconnesso. I dati precedenti sono salvati; avvia una nuova prova.")
                    drift = time.time() - (anchor_epoch + time.monotonic() - anchor_mono)
                    if abs(drift) > 0.5:
                        raise RuntimeError("L'orologio PC e' cambiato di oltre 0.5 s. Registrazione interrotta per proteggere i timestamp.")
                    sample = nearest_microbit(buffer, target, args.tolerance)
                    row = {"n": index + 1, "target_epoch": target_epoch,
                           "status": "GARMIN_DA_SCARICARE" if sample else "MICROBIT_MANCANTE"}
                    if sample:
                        row.update(micro_epoch=sample.epoch, micro_bpm=sample.values["bpm"],
                                   signal=sample.values["signal"], filtered=sample.values["filtered"],
                                   raw=sample.values["raw"])
                    session["rows"].append(row)
                    # Finestre disgiunte: nessun campione puo' essere riutilizzato.
                    while buffer and buffer[0].mono <= target + args.tolerance:
                        buffer.popleft()
                    checkpoint(output, session_path, session)
                    print(f"{index + 1}/{args.samples}  {utc_iso(target_epoch)}  microbit={row.get('micro_bpm', 'mancante')} BPM")
                session["metadata"]["stato_registrazione"] = "completata"
        except (KeyboardInterrupt, asyncio.CancelledError):
            session["metadata"]["stato_registrazione"] = "interrotta dall'utente"
            raise
        except Exception:
            session["metadata"]["stato_registrazione"] = "interrotta per errore"
            raise
        finally:
            excel_saved = checkpoint(output, session_path, session)
            print(f"Sessione salvata: {session_path.name}")
            if excel_saved:
                print(f"Excel salvato: {output.name}")


def match_garmin(session: dict, samples: list[tuple[float, int]], tolerance: float,
                 offset: float = 0.0) -> None:
    """Confronto dei timestamp, senza media, interpolazione o riempimento dei buchi."""
    ordered = sorted(samples)
    times = [epoch + offset for epoch, _ in ordered]
    used = set()
    for row in session["rows"]:
        for key in ("garmin_original_epoch", "garmin_epoch", "garmin_bpm", "delta_seconds"):
            row.pop(key, None)
        micro_epoch = row.get("micro_epoch")
        if micro_epoch is None:
            row["status"] = "MICROBIT_MANCANTE"
            continue
        # Richiedi vicinanza sia alla griglia, sia al campione micro:bit effettivo.
        low = max(micro_epoch, row["target_epoch"]) - tolerance
        high = min(micro_epoch, row["target_epoch"]) + tolerance
        begin, end = bisect.bisect_left(times, low), bisect.bisect_right(times, high)
        candidates = [i for i in range(begin, end) if i not in used]
        best = min(candidates, key=lambda i: (abs(times[i] - micro_epoch), times[i]), default=None)
        if best is None:
            row["status"] = "GARMIN_MANCANTE_ENTRO_TOLLERANZA"
            continue
        used.add(best)
        row.update(garmin_original_epoch=ordered[best][0], garmin_epoch=times[best],
                   garmin_bpm=ordered[best][1], delta_seconds=round(times[best] - micro_epoch, 6),
                   status="ABBINATO_ENTRO_TOLLERANZA")
    session["metadata"].update(tolleranza_abbinamento_secondi=tolerance,
                               correzione_orologio_garmin_secondi=offset)


def load_session(path: Path) -> dict:
    session = json.loads(path.read_text(encoding="utf-8"))
    if session.get("schema_version") != 1 or not isinstance(session.get("rows"), list):
        raise ValueError("Formato del file sessione non supportato.")
    return session


def merge(args) -> None:
    from garmin_api import fetch_activity_fit, read_fit_samples

    session = load_session(args.session)
    output = args.output or args.session.with_name(args.session.name.removesuffix(".session.json") + ".xlsx")
    if args.fit:
        data = args.fit.read_bytes()
        origin = f"FIT locale: {args.fit.name}"
    else:
        cache = args.session.with_name(f"garmin_activity_{args.activity_id}.fit")
        data = fetch_activity_fit(args.activity_id, args.tokens, cache, args.email)
        origin = f"Garmin Connect API (client non ufficiale), attivita' {args.activity_id}"
    samples = read_fit_samples(data)
    match_garmin(session, samples, args.tolerance, args.garmin_offset)
    session["metadata"].update(origine_garmin=origin, download_o_importazione_utc=utc_iso(time.time()))
    excel_saved = checkpoint(output.resolve(), args.session, session)
    if not any(row["status"] == "ABBINATO_ENTRO_TOLLERANZA" for row in session["rows"]):
        raise RuntimeError("Nessuna coppia abbinata. Controlla attivita', orologi e registrazione a 1 secondo; non aumentare la tolleranza per nascondere un errore di orario.")
    if not excel_saved:
        raise RuntimeError("Abbinamento salvato nel JSON, ma Excel non e' stato aggiornato. Chiudi Excel e usa export.")
    print(f"Excel aggiornato: {output.resolve()}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    scan = sub.add_parser("scan", help="Elenca dispositivi Bluetooth visibili")
    scan.add_argument("--scan-seconds", type=float, default=8)
    rec = sub.add_parser("record", help="Registra micro:bit ogni 10 s; Garmin viene aggiunto dopo")
    rec.add_argument("--output", type=Path, default=Path("misure.xlsx"))
    rec.add_argument("--samples", type=int, default=40)
    rec.add_argument("--warmup", type=float, default=20)
    rec.add_argument("--tolerance", type=float, default=1)
    rec.add_argument("--name", default="BBC micro:bit")
    rec.add_argument("--address")
    rec.add_argument("--characteristic")
    rec.add_argument("--pair", action="store_true")
    rec.add_argument("--framing", choices=["packet", "line"], default="packet")
    rec.add_argument("--scan-seconds", type=float, default=8)
    mer = sub.add_parser("merge", help="Scarica l'attivita' Garmin e abbina i BPM")
    mer.add_argument("--session", type=Path, required=True)
    source = mer.add_mutually_exclusive_group(required=True)
    source.add_argument("--activity-id", help="ID numerico dell'attivita' Garmin Connect sincronizzata")
    source.add_argument("--fit", type=Path, help="Alternativa: FIT originale gia' scaricato")
    mer.add_argument("--tokens", type=Path, default=Path(".garmin_tokens"))
    mer.add_argument("--email")
    mer.add_argument("--tolerance", type=float, default=1)
    mer.add_argument("--garmin-offset", type=float, default=0,
                     help="Correzione nota in secondi DA AGGIUNGERE all'orario Garmin; default 0")
    mer.add_argument("--output", type=Path)
    exp = sub.add_parser("export", help="Ricrea Excel dalla sessione, ad esempio se Excel era aperto")
    exp.add_argument("--session", type=Path, required=True)
    exp.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.command in ("record", "merge"):
        if not math.isfinite(args.tolerance) or not 0 <= args.tolerance < INTERVAL / 2:
            parser.error("--tolerance deve essere >= 0 e < 5 s (finestre senza sovrapposizione).")
    if args.command == "record" and (args.samples < 1 or not math.isfinite(args.warmup) or args.warmup < 0):
        parser.error("--samples deve essere positivo e --warmup non negativo e finito.")
    if args.command in ("record", "scan") and (not math.isfinite(args.scan_seconds) or args.scan_seconds <= 0):
        parser.error("--scan-seconds deve essere positivo e finito.")
    if args.command == "merge":
        if not math.isfinite(args.garmin_offset):
            parser.error("--garmin-offset deve essere finito.")
        if args.activity_id and not re.fullmatch(r"[0-9]+", args.activity_id):
            parser.error("--activity-id deve contenere solo cifre.")
    try:
        if args.command == "scan":
            asyncio.run(discover(args))
        elif args.command == "record":
            asyncio.run(record(args))
        elif args.command == "merge":
            merge(args)
        else:
            write_xlsx(args.output, load_session(args.session))
            print(f"Excel salvato: {args.output.resolve()}")
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f"Errore: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
