# -*- coding: utf-8 -*-
"""
엑셀 FILTER 검색 함수식 생성기 (Streamlit 웹앱)

실행 방법
    pip install streamlit
    streamlit run excel_filter_builder.py

입력 규칙
    ,  : 여러 TERM을 구분 (같은 항목 안에서는 OR, 제외 TERM은 옵션에 따라 AND/OR)
    *  : 한 TERM 안의 와일드카드 (예: 금융*제도*개편)
    경로: 엑셀 범위 (예: D2:D1000, Sheet1!D2:D1000)
          쉼표로 여러 범위를 적으면 이어 붙여서 검색 (예: D2:D1000, E2:E1000)
"""
from datetime import datetime

import streamlit as st


# ─────────────────────────── 함수식 생성 로직 ───────────────────────────
def split_csv(text: str) -> list[str]:
    """쉼표로 구분된 문자열을 리스트로 변환한다."""
    return [t.strip() for t in (text or "").replace("，", ",").split(",") if t.strip()]


def quote(text: str) -> str:
    """엑셀 문자열 리터럴로 변환한다 (큰따옴표 이스케이프)."""
    return '"' + text.replace('"', '""') + '"'


def range_expr(path: str) -> str:
    """경로 입력을 엑셀 범위식으로 변환한다. 여러 범위는 공백을 두고 이어 붙인다."""
    parts = split_csv(path)
    return '&" "&'.join(parts)


def term_cond(term: str, rng: str, star_mode: str) -> str:
    """TERM 하나에 대한 조건식을 만든다."""
    if star_mode == "AND":
        pieces = [p.strip() for p in term.split("*") if p.strip()]
        conds = [f"ISNUMBER(SEARCH({quote(p)},{rng}))" for p in pieces]
        return conds[0] if len(conds) == 1 else "(" + "*".join(conds) + ")"
    return f"ISNUMBER(SEARCH({quote(term)},{rng}))"


def or_group(conds: list[str]) -> str:
    return conds[0] if len(conds) == 1 else "((" + "+".join(conds) + ")>0)"


def parse_date(text: str):
    """YYYYMMDD 문자열을 date로 변환한다. 형식이 틀리면 None."""
    try:
        return datetime.strptime(text.strip(), "%Y%m%d").date()
    except ValueError:
        return None


def build(v: dict) -> tuple[str, str, list[str]]:
    """입력값 dict를 받아 (Query Preview, 함수식, 오류목록)을 반환한다."""
    errors: list[str] = []
    conds: list[str] = []          # 함수식 조건 (서로 AND)
    kw_preview: list[str] = []     # Preview 1행 (검색어)
    tail_preview: list[str] = []   # Preview 2행 이후 (서지사항)

    def need_path(label: str, path: str) -> str:
        rng = range_expr(path)
        if not rng:
            errors.append(f"{label}의 경로가 비어 있습니다.")
        return rng

    star = v["star_mode"]

    # (주 OR 관련어)
    main_conds, main_terms = [], []
    for label, key in (("TERM(주)", "main"), ("TERM(관련어)", "rel")):
        terms = split_csv(v[key])
        if terms:
            rng = need_path(label, v[key + "_path"])
            main_conds += [term_cond(t, rng, star) for t in terms]
            main_terms += terms
    if main_conds:
        conds.append(or_group(main_conds))
        kw_preview.append("(" + " OR ".join(main_terms) + ")")

    # (부1 OR 부2 ...)
    sub_terms = split_csv(v["sub"])
    if sub_terms:
        rng = need_path("TERM(부)", v["sub_path"])
        conds.append(or_group([term_cond(t, rng, star) for t in sub_terms]))
        kw_preview.append("(" + " OR ".join(sub_terms) + ")")

    # NOT (제외1 AND/OR 제외2 ...)
    exc_terms = split_csv(v["exc"])
    if exc_terms:
        rng = need_path("TERM(제외)", v["exc_path"])
        exc_conds = [term_cond(t, rng, star) for t in exc_terms]
        if len(exc_conds) == 1:
            conds.append(f"NOT({exc_conds[0]})")
        else:
            op = "*" if v["exc_mode"] == "AND" else "+"
            conds.append("((" + op.join(exc_conds) + ")=0)")
        kw_preview.append("NOT (" + f" {v['exc_mode']} ".join(exc_terms) + ")")

    # 발간처 (완전 일치, OR)
    pubs = split_csv(v["pub"])
    if pubs:
        rng = need_path("발간처", v["pub_path"])
        conds.append(or_group([f"({rng}={quote(p)})" for p in pubs]))
        tail_preview.append("(발간처=" + " OR ".join(pubs) + ")")

    # 발간일 (입력일 이후)
    if v["date"].strip():
        d = parse_date(v["date"])
        rng = split_csv(v["date_path"])
        if d is None:
            errors.append("발간일은 YYYYMMDD 형식의 유효한 날짜여야 합니다.")
        elif not rng:
            errors.append("발간일의 경로가 비어 있습니다.")
        else:
            if v["date_type"] == "날짜":
                conds.append(f"({rng[0]}>=DATE({d.year},{d.month},{d.day}))")
            else:  # 셀 값이 YYYYMMDD 숫자/텍스트인 경우
                conds.append(f'({rng[0]}&"">="{d:%Y%m%d}")')
            tail_preview.append(f"(DATE>={d:%Y,%m,%d})")

    ret = v["ret_path"].strip()
    if not ret:
        errors.append("결과 반환 범위가 비어 있습니다.")
    if not conds and not errors:
        errors.append("조건을 하나 이상 입력하세요.")

    lines = []
    if kw_preview:
        lines.append("(" + " AND ".join(kw_preview) + ")")
    lines += tail_preview
    preview = "\nAND ".join(lines)

    sep = "\n *" if v["multiline"] else "*"
    if v["multiline"]:
        formula = f"=FILTER({ret},\n  " + sep.join(conds) + ',\n  "결과 없음")'
    else:
        formula = f"=FILTER({ret}," + sep.join(conds) + ',"결과 없음")'
    return preview, formula, errors


# ─────────────────────────────── 화면 ───────────────────────────────
def field(label: str, key: str, value: str, path: str, placeholder: str = "") -> None:
    """[값 입력 | 경로 입력] 한 줄을 그린다."""
    c1, c2 = st.columns([3, 2])
    c1.text_input(label, value=value, key=key, placeholder=placeholder)
    c2.text_input(f"{label} 경로", value=path, key=key + "_path", placeholder="예: D2:D1000")


def main() -> None:
    st.set_page_config(page_title="엑셀 검색 함수식 생성기", layout="wide")
    st.title("엑셀 검색 함수식 생성기")
    st.caption("TERM은 쉼표(,)로 구분하고, 한 TERM 안에서는 *를 사용합니다. "
               "경로에 범위를 쉼표로 여러 개 적으면 이어 붙여 검색합니다. 모든 범위는 행 수가 같아야 합니다.")

    st.subheader("서지사항")
    field("발간처", "pub", "기획예산처, 금융위원회, 국토교통부", "B2:B1000")
    field("발간일 (이후)", "date", "20260101", "C2:C1000", "YYYYMMDD")

    st.subheader("검색어 입력")
    field("TERM(주)", "main", "금융*제도*개편", "D2:D1000")
    field("TERM(관련어)", "rel", "파이낸스*정책*개선", "D2:D1000")
    field("TERM(부)", "sub", "혁신,보안,AI,기술개발,리스크*관리", "D2:D1000")
    field("TERM(제외)", "exc", "협력,공급망", "D2:D1000")

    st.subheader("출력 설정")
    c1, c2, c3, c4 = st.columns(4)
    c1.text_input("결과 반환 범위", value="A2:D1000", key="ret_path")
    c2.selectbox("발간일 셀 형식", ["날짜", "YYYYMMDD"], key="date_type")
    c3.selectbox("제외 TERM 결합", ["AND", "OR"], key="exc_mode",
                 help="AND: 제외어가 모두 있는 행만 제외 / OR: 하나라도 있으면 제외")
    c4.selectbox("* 처리", ["와일드카드", "AND"], key="star_mode",
                 help="와일드카드: 금융…제도…개편 순서대로 등장 / AND: 순서와 무관하게 모두 포함")
    st.checkbox("함수식 줄바꿈 표시", key="multiline")

    keys = ["pub", "date", "main", "rel", "sub", "exc"]
    values = {k: st.session_state[k] for k in keys}
    values |= {k + "_path": st.session_state[k + "_path"] for k in keys}
    values |= {k: st.session_state[k]
               for k in ("ret_path", "date_type", "exc_mode", "star_mode", "multiline")}

    preview, formula, errors = build(values)

    st.subheader("연산식 Query Preview")
    st.code(preview or " ", language=None)

    st.subheader("함수식")
    if errors:
        for e in errors:
            st.error(e)
    else:
        st.code(formula, language="excel")
        if len(formula) > 8192:
            st.warning(f"함수식 길이({len(formula):,}자)가 엑셀 한도(8,192자)를 넘습니다.")


if __name__ == "__main__":
    main()
