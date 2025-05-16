import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import re
from io import BytesIO
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="DR 자동 생성기", layout="wide")
st.title("📦 DR.XLSX 자동 생성 프로그램")

# ──────────────────────────────────────────────
# 유틸리티
# ──────────────────────────────────────────────
def clean_text(text: str) -> str:
    """특정 불용어·기호 제거 후 공백 제거"""
    if not isinstance(text, str):
        return ""
    patterns = [
        r"#.*?セット", r"【.*?】", r"/.*?", r"韓コスメ", r"口紅", r"リップ", r"アワグロウ",
        r"[\[\]【】#]", r"\s{2,}"
    ]
    for pat in patterns:
        text = re.sub(pat, "", text)
    return re.sub(r"\s+", "", text.strip())


def match_items(src_list, tgt_list, threshold=0.3):
    """SequenceMatcher 기반 유사도 매핑"""
    mapping = {}
    for i, src in enumerate(src_list):
        best, best_idx = 0, None
        for j, tgt in enumerate(tgt_list):
            score = SequenceMatcher(None, src, tgt).ratio()
            if score > threshold and score > best:
                best, best_idx = score, j
        if best_idx is not None:
            mapping[i] = best_idx
    return mapping


def format_postal(postal):
    """
    입력:
      - 6·7자리 숫자  →  3-4 형식
      - 이미 3-4 형식 → 그대로 유지
      - 그 외         → '우편번호없음'
    """
    if pd.isna(postal):
        return "우편번호없음"

    # 이미 xxx-xxxx 형태인가?
    if isinstance(postal, str) and re.fullmatch(r"\d{3}-\d{4}", postal.strip()):
        return postal.strip()

    # 숫자·float → 문자 변환
    if isinstance(postal, (int, float)):
        postal = str(int(postal))
    if isinstance(postal, str):
        digits = postal.zfill(7)           # 6자리면 앞에 0
        if digits.isdigit() and len(digits) == 7:
            return f"{digits[:3]}-{digits[3:]}"
    return "우편번호없음"


# ──────────────────────────────────────────────
# 기본 H 로더 (캐시)
# ──────────────────────────────────────────────
@st.cache_data
def load_default_h():
    default_path = Path(__file__).with_name("H.xlsx")
    if default_path.exists():
        return pd.read_excel(default_path)
    else:
        st.error("⚠️  기본 H.xlsx 가 앱 폴더에 없습니다.")
        return pd.DataFrame()

# ──────────────────────────────────────────────
# 사이드바 – H 파일 선택
# ──────────────────────────────────────────────
st.sidebar.header("H.XLSX 관리")
use_default_h = st.sidebar.checkbox("기본 H.XLSX 사용", value=True)

if use_default_h:
    df_H = load_default_h()
else:
    h_file = st.sidebar.file_uploader("H.XLSX 업로드", type=["xlsx"])
    if h_file:
        df_H = pd.read_excel(h_file)
    else:
        st.warning("H.XLSX 파일을 업로드하거나 기본 파일을 사용해 주세요.")
        st.stop()

# 컬럼명 소문자 통일
df_H.columns = df_H.columns.str.lower()

# ──────────────────────────────────────────────
# 본문 – S 파일 업로드
# ──────────────────────────────────────────────
st.subheader("1단계 · S.XLSX 업로드")
s_file = st.file_uploader("S.XLSX 파일을 선택하세요", type=["xlsx"])

if s_file:
    df_S = pd.read_excel(s_file)
    df_S.columns = df_S.columns.str.lower()

    # 1) S ↔ H 매핑
    s_names_clean = [clean_text(x) for x in df_S["item_name"].fillna("")]
    h_names_clean = [clean_text(x) for x in df_H["출고상품명"].fillna("")]
    s2h = match_items(s_names_clean, h_names_clean)

    df_S_upd = df_S.copy()
    for s_idx, h_idx in s2h.items():
        df_S_upd.at[s_idx, "상품 shoppingmall url"] = df_H.at[h_idx, "상품 shoppingmall url"]
        df_S_upd.at[s_idx, "unit_total price"]    = df_H.at[h_idx, "unit_total price"]

    # 2) 규칙 적용
    df_S_upd["order_no"] = df_S_upd["order_no"].astype(str).apply(
        lambda x: "86" + x if not x.startswith("86") else x
    )
    for col, val in [
        ("service code",            "99"),
        ("consignee_국가코드",       "JP"),
        ("pkg",                     "1"),
        ("item_origin",             "KR"),
        ("currency",                "JPY"),
    ]:
        df_S_upd[col] = df_S_upd.apply(
            lambda r, c=col, v=val: v if r.dropna().shape[0] > 1 else r[c], axis=1
        )

    df_S_upd["consignee_address (en)_jp지역 현지어 기재"] = df_S_upd[
        "consignee_address (en)_jp지역 현지어 기재"
    ].apply(lambda x: re.sub(r"\[.*?\]", "", x) if isinstance(x, str) else x)

    # ⬇️  **우편번호 로직 강화** (xxx-xxxx 형태 유지)
    df_S_upd["consignee_ postalcode"] = df_S_upd["consignee_ postalcode"].apply(format_postal)

    # 3) DR 시트 생성
    dr_cols = ["ref_no (주문번호)", "하이브 상품코드", "상품명", "수량", "바코드"]
    df_DR = pd.DataFrame(columns=dr_cols)
    df_DR["ref_no (주문번호)"] = df_S_upd["order_no"]
    df_DR["상품명"]            = df_S_upd["item_name"]
    df_DR["수량"]              = df_S_upd["item_pcs"]

    dr2h = match_items([clean_text(x) for x in df_DR["상품명"].fillna("")], h_names_clean)
    for dr_idx, h_idx in dr2h.items():
        df_DR.at[dr_idx, "하이브 상품코드"] = df_H.at[h_idx, "상품코드"]
        df_DR.at[dr_idx, "바코드"]        = df_H.at[h_idx, "바코드"]
        df_DR.at[dr_idx, "상품명"]        = df_H.at[h_idx, "출고상품명"]

    st.success("🎉 DR 파일 생성 완료!")
    st.dataframe(df_DR.head())

    # 4) 다운로드 버튼
    today = datetime.today().strftime("%y%m%d")
    buf_s, buf_dr = BytesIO(), BytesIO()
    df_S_upd.to_excel(buf_s,  index=False); buf_s.seek(0)
    df_DR.to_excel(buf_dr,    index=False); buf_dr.seek(0)

    st.download_button(
        "📥 RINCOS_온드_주문등록양식_큐텐.xlsx",
        buf_s,
        file_name=f"{today}_RINCOS_온드_주문등록양식_큐텐.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "📥 RINCOS_온드_HIVE센터 B2C 출고요청양식.xlsx",
        buf_dr,
        file_name=f"{today}_RINCOS_온드_HIVE센터 B2C 출고요청양식.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("좌측에서 H.XLSX 설정 후 S.XLSX 파일을 업로드하세요.")
