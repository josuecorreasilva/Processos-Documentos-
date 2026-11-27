#!/usr/bin/env python3
"""Gera e valida programaticamente os 3 documentos .docx do Processo 3 — PICSPel Materiais de Consumo."""

from decimal import Decimal
from pathlib import Path
import re
import sys

from docx import Document

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.picspel_consumo.dados import (  # noqa: E402
    ITENS_CONSUMO,
    META,
    TOTAL_CONSUMO,
    TOTAL_UNIDADES,
    fmt_brl,
    fmt_int,
)
from scripts.picspel_consumo.gerar_dfd import gerar_dfd  # noqa: E402
from scripts.picspel_consumo.gerar_etp import gerar_etp  # noqa: E402
from scripts.picspel_consumo.gerar_tr import gerar_tr  # noqa: E402

OUT_DIR = REPO_ROOT / "docs" / "10-PICSPEL-MATERIAIS-CONSUMO"


def parse_brl(text: str) -> Decimal:
    cleaned = text.replace("R$", "").strip().replace(".", "").replace(",", ".")
    return Decimal(cleaned)


def full_text(doc: Document) -> str:
    parts = [p.text for p in doc.paragraphs]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def validate_item_table(table, doc_label: str, *, catmat_col: bool) -> None:
    # 1 header + 15 items + 1 total row = 17 rows
    assert len(table.rows) == 17, f"{doc_label}: esperadas 17 linhas na tabela de itens, obtidas {len(table.rows)}"
    qtd_col = 4 if catmat_col else 3
    unit_col = 5 if catmat_col else 4
    tot_col = 6 if catmat_col else 5

    soma_calc = Decimal("0.00")
    soma_qtd = 0
    for idx, item in enumerate(ITENS_CONSUMO, start=1):
        row = table.rows[idx]
        cod = row.cells[0].text.strip()
        desc = row.cells[1].text.strip()
        qtd = int(row.cells[qtd_col].text.strip().replace(".", ""))
        v_unit = parse_brl(row.cells[unit_col].text)
        v_tot = parse_brl(row.cells[tot_col].text)

        assert cod == item.codigo, f"{doc_label} linha {idx}: código {cod} != {item.codigo}"
        assert desc == item.especificacao, f"{doc_label} item {cod}: especificação divergente da fonte única"
        assert qtd == item.quantidade, f"{doc_label} item {cod}: qtd {qtd} != {item.quantidade}"
        assert v_unit == item.valor_unitario, f"{doc_label} item {cod}: v_unit {v_unit} != {item.valor_unitario}"
        assert v_tot == item.valor_total, f"{doc_label} item {cod}: v_tot {v_tot} != {item.valor_total}"
        assert v_unit * qtd == v_tot, f"{doc_label} item {cod}: aritmética unit × qtd != total"

        soma_calc += v_tot
        soma_qtd += qtd

    total_row = table.rows[-1]
    tot_declared = parse_brl(total_row.cells[-1].text)
    qtd_declared = int(total_row.cells[qtd_col].text.strip().replace(".", ""))

    assert soma_calc == TOTAL_CONSUMO == tot_declared, (
        f"{doc_label}: soma calculada {soma_calc} != declarada {tot_declared}"
    )
    assert soma_qtd == TOTAL_UNIDADES == qtd_declared, (
        f"{doc_label}: qtd somada {soma_qtd} != declarada {qtd_declared}"
    )


def validate_header_and_signatures(doc: Document, doc_label: str) -> None:
    hdr = doc.sections[0].header
    hdr_text = "\n".join(p.text for p in hdr.paragraphs)
    assert "MUNICÍPIO DE PELOTAS" in hdr_text, f"{doc_label}: cabeçalho sem MUNICÍPIO DE PELOTAS"
    assert "SECRETARIA MUNICIPAL DE SAÚDE" in hdr_text, f"{doc_label}: cabeçalho sem SMS"
    # Verifica presença da imagem do brasão no cabeçalho (blip)
    hdr_xml = hdr._element.xml
    assert "a:blip" in hdr_xml, f"{doc_label}: brasão ausente no cabeçalho"

    ftr_xml = doc.sections[0].footer._element.xml
    assert "PAGE" in ftr_xml and "NUMPAGES" in ftr_xml, f"{doc_label}: paginação Página X de Y ausente"

    txt = full_text(doc)
    for required in (
        META["gestor_nome"],
        META["gestor_matricula"],
        META["fiscal_nome"],
        META["fiscal_matricula"],
        META["secretaria_nome"],
        META["emenda"],
        "3.3.90.30",
        fmt_brl(TOTAL_CONSUMO),
    ):
        assert required in txt, f"{doc_label}: texto obrigatório ausente: {required}"

    # Verifica ausência de marca comercial 'micropore' e de resquícios do modelo RP05
    for forbidden in ("micropore", "Vias Aéreas", "Cânulas Guedel", "TUBO ENDOTRAQUEAL", "Josué Corrêa"):
        assert forbidden.lower() not in txt.lower(), f"{doc_label}: termo proibido encontrado: {forbidden}"


def validate_all(dfd_path: Path, etp_path: Path, tr_path: Path) -> None:
    dfd = Document(str(dfd_path))
    etp = Document(str(etp_path))
    tr = Document(str(tr_path))

    # 1. Validação DFD
    validate_header_and_signatures(dfd, "DFD")
    validate_item_table(dfd.tables[0], "DFD", catmat_col=False)

    # 2. Validação ETP
    validate_header_and_signatures(etp, "ETP")
    validate_item_table(etp.tables[0], "ETP", catmat_col=False)
    # Tabela 2 do ETP (classificação por natureza) -> tabela de índice 2 (índice 1 é o callout)
    nat_table = next(t for t in etp.tables if len(t.rows) == 17 and len(t.columns) == 6 and "ANVISA" in t.rows[0].cells[4].text)
    soma_nat = sum((parse_brl(nat_table.rows[i].cells[5].text) for i in range(1, 16)), Decimal("0.00"))
    tot_nat = parse_brl(nat_table.rows[-1].cells[-1].text)
    assert soma_nat == TOTAL_CONSUMO == tot_nat, f"ETP tabela natureza: {soma_nat} != {tot_nat}"

    # 3. Validação TR (exatas 14 seções do modelo RP05_TR)
    validate_header_and_signatures(tr, "TR")
    validate_item_table(tr.tables[0], "TR", catmat_col=True)
    tr_text = full_text(tr)
    expected_tr_sections = [
        "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO",
        "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO",
        "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO",
        "4 GESTÃO/ANÁLISE DE RISCOS",
        "5 REQUISITOS DA CONTRATAÇÃO",
        "6 DA EXECUÇÃO DO OBJETO",
        "7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO",
        "8 RECEBIMENTO DO OBJETO",
        "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO",
        "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR",
        "11 DAS OBRIGAÇÕES DO CONTRATANTE",
        "12 DAS OBRIGAÇÕES DO CONTRATADO",
        "13 DA ADEQUAÇÃO ORÇAMENTÁRIA",
        "14 EQUIPE DE PLANEJAMENTO",
    ]
    for sec_title in expected_tr_sections:
        assert sec_title in tr_text, f"TR: seção obrigatória ausente: {sec_title}"

    # Confirma que 47.200 só aparece como nota histórica no DFD/ETP e nunca no TR
    assert "47.200" not in tr_text, "TR não deve conter o valor nominal incorreto 47.200"
    for label, txt in (("DFD", full_text(dfd)), ("ETP", full_text(etp))):
        # Em DFD e ETP, 47.200 só pode aparecer 1 vez na nota de revisão aritmética
        matches = re.findall(r"47\.200", txt)
        assert len(matches) == 1, f"{label}: esperado 1 registro histórico de 47.200, encontrados {len(matches)}"

    print("VALIDAÇÃO PROGRAMÁTICA CONCLUÍDA COM SUCESSO!")
    print(f"  - Itens conferidos (DFD, ETP e TR): {len(ITENS_CONSUMO)} itens ({fmt_int(TOTAL_UNIDADES)} unidades comerciais)")
    print(f"  - Soma aritmética exata validada em todas as tabelas: {fmt_brl(TOTAL_CONSUMO)}")
    print(f"  - DFD: {dfd_path} ({dfd_path.stat().st_size:,} bytes)")
    print(f"  - ETP: {etp_path} ({etp_path.stat().st_size:,} bytes)")
    print(f"  - TR : {tr_path} ({tr_path.stat().st_size:,} bytes)")


def main() -> None:
    dfd_path = gerar_dfd(OUT_DIR / "PICSPel_Materiais_Consumo_DFD_Revisado.docx")
    etp_path = gerar_etp(OUT_DIR / "PICSPel_Materiais_Consumo_ETP_Revisado.docx")
    tr_path = gerar_tr(OUT_DIR / "PICSPel_Materiais_Consumo_TR.docx")
    validate_all(dfd_path, etp_path, tr_path)


if __name__ == "__main__":
    main()
