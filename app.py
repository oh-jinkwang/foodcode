import pandas as pd
import streamlit as st

st.set_page_config(page_title="식품/위생용품/첨가물 기준 및 규격 검색", page_icon="??")

DATA_FILE = "standards.xlsx"
SHEET_DEF, SHEET_COMMON, SHEET_INDIV = "정의", "공통기준", "개별기준"

# 식용유 계열을 검색하면 '식물성유지류' 공통기준도 함께 표시
OIL_PARENT = "식물성유지류"
OIL_SUBTYPES = {
    "콩기름", "옥수수기름", "채종유", "미강유", "참기름", "추출참깨유", "들기름", "추출들깨유",
    "홍화유", "해바라기유", "목화씨기름", "땅콩기름", "올리브유", "팜유", "야자유",
    "고추씨기름", "기타식물성유지",
}

# 시트의 열 위치 (A=0, B=1, ...) - 기존 앱스크립트와 동일
COL_DEF = {"type": 0, "text": 1, "updated": 2}
COL_COMMON = {"self_check": 0, "item": 1, "type": 3, "proviso": 4, "unit": 6, "summary": 7}
COL_INDIV = {"self_check": 0, "item": 1, "type": 2, "proviso": 3, "unit": 5,
             "summary": 6, "detail": 7, "start": 8}

COMMON_VIEW = {"type": "식품유형", "item": "시험항목명", "summary": "기준규격요약",
               "unit": "단위명", "proviso": "단서조항명", "self_check": "자가품질검사"}
INDIV_VIEW = {"type": "식품유형", "item": "시험항목명", "detail": "세부항목명",
              "summary": "기준규격요약", "unit": "단위명", "proviso": "단서조항명",
              "self_check": "자가품질검사"}
QUOTES = [('"', '"'), ("“", "”"), ("'", "'"), ("‘", "’")]


@st.cache_data(show_spinner="기준 데이터를 불러오는 중...")
def load_data():
    xls = pd.ExcelFile(DATA_FILE)

    def read(name, mapping):
        if name not in xls.sheet_names:
            raise ValueError(f"'{name}' 탭이 없습니다. 엑셀의 탭 이름: {xls.sheet_names}")
        raw = pd.read_excel(xls, sheet_name=name, dtype=str).fillna("")
        raw = raw.apply(lambda c: c.astype(str).str.strip())
        need = max(mapping.values()) + 1
        if raw.shape[1] < need:
            raise ValueError(f"'{name}' 탭의 열이 {raw.shape[1]}개뿐입니다 (최소 {need}개 필요).")
        df = pd.DataFrame({k: raw.iloc[:, v] for k, v in mapping.items()})
        return df[df["type"] != ""]

    # 정의: 식품유형별로 최종 업데이트 날짜가 가장 최신인 1건만 사용
    d = read(SHEET_DEF, COL_DEF)
    d["dt"] = pd.to_datetime(d["updated"].str.replace(".", "-", regex=False), errors="coerce")
    d = d.sort_values("dt", na_position="first", kind="stable").drop_duplicates("type", keep="last")
    def_map = d.set_index("type")["text"].to_dict()

    common = read(SHEET_COMMON, COL_COMMON)

    # 개별기준: 7개 항목이 모두 같으면 유효개시일(YYYYMMDD)이 가장 최신인 1건만 남김
    indiv = read(SHEET_INDIV, COL_INDIV)
    indiv["start_n"] = indiv["start"].str.replace(r"\D", "", regex=True).str[:8]
    key = ["type", "item", "detail", "summary", "unit", "proviso", "self_check"]
    indiv = (indiv.sort_values("start_n", kind="stable")
                  .drop_duplicates(key, keep="last").sort_index())
    return def_map, common, indiv


def parse_query(raw, exact_checked):
    q, exact = raw.strip(), exact_checked
    for o, c in QUOTES:                      # "두부" 처럼 따옴표로 감싸면 정확일치
        if len(q) >= 2 and q[0] == o and q[-1] == c:
            return q[1:-1].strip(), True
    return q, exact


@st.cache_data(show_spinner=False)           # 같은 검색은 즉시 표시
def run_search(query, exact):
    def_map, common, indiv = load_data()
    match = (lambda s: s == query) if exact else (lambda s: s.str.contains(query, regex=False))

    def_types = [t for t in def_map if (t == query if exact else query in t)]
    c = common[match(common["type"])].copy()
    i = indiv[match(indiv["type"])].copy()
    types = set(def_types) | set(c["type"]) | set(i["type"])

    if types & OIL_SUBTYPES:
        parent = common[common["type"] == OIL_PARENT]
        if not parent.empty:
            c = pd.concat([c, parent]).drop_duplicates(["type", "item", "summary"])
            types.add(OIL_PARENT)

    types = sorted(types)                    # 가나다순
    c = c.sort_values("type", kind="stable")
    i = i.sort_values("type", kind="stable")
    return types, {t: def_map.get(t, "") for t in types}, c, i


def show_table(df, view):
    out = df[list(view)].rename(columns=view).reset_index(drop=True)
    styled = out.style.apply(                # 자가품질검사 Y 행은 노란색
        lambda row: ["background-color: #FFF6D2"] * len(row)
        if row["자가품질검사"].upper() == "Y" else [""] * len(row), axis=1)
    st.dataframe(styled, hide_index=True)


# ------------------------------ 화면 ------------------------------
st.title("식품/위생용품/첨가물 기준 및 규격 검색")
st.markdown("유형을 입력하면 정의, 관련기준, 규격을 검색합니다.  \n일부 단어만 입력해도 검색 가능 합니다.")

with st.form("search_form", clear_on_submit=True):   # 엔터로도 검색, 검색 후 입력창 비움
    raw = st.text_input("식품유형", label_visibility="collapsed",
                        placeholder="ex) 두부, 된장 (일부만 입력해도 검색)")
    submitted = st.form_submit_button("찾기", type="primary")
exact_chk = st.checkbox("정확히 일치하는 식품유형만 검색")

if submitted:
    if raw.strip():
        st.session_state["query"] = raw.strip()
    else:
        st.warning("식품유형을 입력해주세요.")

if st.session_state.get("query"):
    q, exact = parse_query(st.session_state["query"], exact_chk)
    try:
        types, defs, common, indiv = run_search(q, exact)
    except Exception as e:
        st.error(f"데이터를 불러오지 못했습니다: {e}")
        st.stop()

    if not types:
        st.warning(f'"{q}"에 해당하는 식품유형이 없습니다. 철자를 확인해주세요.')
    else:
        with st.container(border=True):
            mode = "정확일치" if exact else "포함"
            st.markdown(f'**"{q}" {mode} 식품유형 {len(types)}건**: {", ".join(types)}')
            for t in types:
                st.markdown(f"**{t}**")
                st.markdown(defs[t] or "-")
        if not common.empty:
            st.subheader(f"공통기준 ({len(common)})")
            show_table(common, COMMON_VIEW)
        if not indiv.empty:
            st.subheader(f"개별기준 ({len(indiv)})")
            show_table(indiv, INDIV_VIEW)
        if common.empty and indiv.empty:
            st.info("해당 식품유형에 대한 공통기준·개별기준 데이터를 찾을 수 없습니다.")