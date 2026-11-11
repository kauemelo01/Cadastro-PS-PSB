"""Validação e aplicação atômica de uma fila offline em uma cópia do XLSX."""

import threading
import uuid
from datetime import datetime

LOG_SHEET = "_PS_SYNC_LOG"
SYNC_LOCK = threading.Lock()
LOG_HEADERS = ("ID DA ALTERAÇÃO", "NUMERO ORIGINAL", "NUMERO GRAVADO", "EDITOR",
               "DATA/HORA ORIGINAL", "CAMPO", "ANTES", "DEPOIS")


def validate_operations(operations, allowed_columns, convert_value, local_tz):
    """Não confia no conteúdo recebido do navegador; mantém a ordem de confirmação."""
    if not isinstance(operations, list) or not operations:
        raise ValueError("Nenhuma alteração para enviar.")
    result, ids = [], set()
    for op in operations:
        op_id = str(uuid.UUID(op["id"]))
        if op_id in ids:
            raise ValueError("ID de alteração repetido no lote.")
        ids.add(op_id)
        numero = str(op["numero"]).strip()
        editor = " ".join(str(op["editor"]).split())
        if not numero or len(editor) < 2:
            raise ValueError("Número e nome de quem alterou são obrigatórios.")
        ts = datetime.fromisoformat(op["timestamp"].replace("Z", "+00:00"))
        if ts.tzinfo is None:
            raise ValueError("A data/hora original deve incluir o fuso horário.")
        changes = {}
        for col, ch in op["changes"].items():
            if col not in allowed_columns or col == "NUMERO":
                raise ValueError(f"Campo não editável: {col}.")
            old, new = ch["old"], ch["new"]
            if not isinstance(old, str) or not isinstance(new, str):
                raise ValueError("Valores de alteração inválidos.")
            if col == "NOME" and not new.strip():
                raise ValueError("O nome não pode ficar em branco.")
            if max(len(old), len(new)) > 32767:
                raise ValueError("Um campo excede o limite de texto da planilha.")
            changes[col] = {"old": old, "new": new, "value": convert_value(col, new)}
        if not changes:
            raise ValueError("Alteração sem campos para gravar.")
        row = op.get("row", {})
        if not isinstance(row, dict) or any(not isinstance(v, str) for v in row.values()):
            raise ValueError("Cadastro de origem inválido.")
        result.append({"id": op_id, "numero": numero, "editor": editor,
                       "ts": ts.astimezone(local_tz).replace(tzinfo=None),
                       "changes": changes, "row": row})
    return result


def apply_batch(wb, operations, patch_record):
    """Aplica tudo antes do upload. O diário torna segura a repetição após timeout.

    A aba de auditoria guarda cada confirmação, mesmo se um cadastro foi editado
    várias vezes offline. As colunas 51/52 continuam mostrando a última edição.
    """
    if LOG_SHEET in wb.sheetnames:
        log = wb[LOG_SHEET]
        if tuple(c.value for c in log[1]) != LOG_HEADERS:
            raise ValueError(f"A aba {LOG_SHEET} tem um formato desconhecido.")
    else:
        log = wb.create_sheet(LOG_SHEET)
        log.append(LOG_HEADERS)
        log.sheet_state = "hidden"
    applied = {r[0]: r[2] for r in log.iter_rows(min_row=2, values_only=True)}
    aliases, results = {}, {}
    for op in operations:
        original = op["numero"]
        if op["id"] in applied:
            target = applied[op["id"]]
        else:
            target = aliases.get(original, original)
            err, new_numero = patch_record(wb, target, op["changes"], op["editor"],
                                          op["ts"], op["row"])
            if err:
                raise ValueError(err)
            target = new_numero or target
            for col, ch in op["changes"].items():
                log.append([op["id"], original, target, op["editor"], op["ts"],
                            col, ch["old"], ch["new"]])
                # Texto começando com '=' deve continuar sendo texto no diário.
                for cell in log[log.max_row]:
                    if isinstance(cell.value, str):
                        cell.data_type = "s"
                log.cell(log.max_row, 5).number_format = "dd/mm/yyyy hh:mm:ss.000"
            applied[op["id"]] = target
        aliases[original] = target
        results[op["id"]] = target
    return results
