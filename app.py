import base64
import io
import json
import os
import tempfile
import time
from datetime import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from typing import List
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
import gspread
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN

# 페이지 기본 설정
st.set_page_config(page_title="IT/OT 보안 아키텍트 AI (하이브리드 + 시트 연동)", layout="wide")

# ==========================================
# 1. 팀원 전용 비밀번호 인증 기능
# ==========================================
def check_password() -> bool:
    if st.session_state.get("authenticated", False):
        return True

    st.title("🔒 IT/OT 보안 아키텍트 AI - 접근 제한")
    st.info("사내 담당자 전용 웹 에이전트입니다. 지정된 비밀번호를 입력해 주세요.")
    
    with st.form("login_form"):
        pwd_input = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
        submitted = st.form_submit_button("접속하기", type="primary")
        
        if submitted:
            correct_pwd = os.environ.get("APP_PASSWORD") or st.secrets.get("APP_PASSWORD", "")
            if not correct_pwd:
                st.error("서버에 APP_PASSWORD가 설정되지 않았습니다. Streamlit Secrets 설정을 확인하세요.")
            elif pwd_input == correct_pwd:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("비밀번호가 일치하지 않습니다.")
    return False

if not check_password():
    st.stop()

# ==========================================
# 2. Google Sheets 히스토리 DB 연동 함수
# ==========================================
def get_gsheet_worksheet():
    if "gcp_service_account" not in st.secrets or "GSHEET_ID" not in st.secrets:
        return None
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["GSHEET_ID"])
        ws = sh.sheet1
        if not ws.row_values(1):
            ws.append_row(["Timestamp", "Project_Title", "User_Prompt", "Chat_History_JSON", "Latest_Design_JSON"])
        return ws
    except Exception as e:
        st.error(f"Google Sheets 연결 오류: {e}")
        return None
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)

def save_history_to_sheet(project_title: str, prompt_text: str, chat_history: list, latest_design_json: str):
    ws = get_gsheet_worksheet()
    if ws is None:
        st.warning("⚠️ Google Sheets 연결 설정(Secrets 또는 시트 공유 권한)을 확인해 주세요.")
        return False
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        trimmed_history = json.dumps(chat_history[-6:], ensure_ascii=False)[:45000]
        trimmed_design = latest_design_json[:48000] if latest_design_json else ""
        ws.append_row([now_str, project_title, prompt_text[:500], trimmed_history, trimmed_design])
        load_history_records.clear()  # 저장 직후 사이드바 목록 캐시 즉시 갱신
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return True
    except Exception as e:
        st.error(f"Google Sheets 저장 실패 원인: {e}")
        return False

@st.cache_data(ttl=60)
def load_history_records() -> list:
    if "gcp_service_account" not in st.secrets or "GSHEET_ID" not in st.secrets:
        return []
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["GSHEET_ID"])
        ws = sh.sheet1
        records = ws.get_all_records()
        return list(reversed(records))
    except Exception as e:
        st.sidebar.error(f"시트 조회 오류: {e}")
        return []

# ==========================================
# 3. 데이터 규격 및 시스템 프롬프트 정의
# ==========================================
class ComplianceItem(BaseModel):
    requirement_id: str = Field(description="RFP 요구사항 번호, 표준 조항 또는 세그멘테이션 구간")
    category: str = Field(description="보안 영역 (예: 마이크로 세그멘테이션, 네트워크 분리, 자산 가시성, 원격접속 통제)")
    as_is_gap: str = Field(description="기존 As-Is 구성 또는 식별 자산의 취약점/혼재 현황")
    to_be_control: str = Field(description="To-Be 구성도에 반영된 보안 통제 방안, 솔루션 및 허용 프로토콜")
    justification: str = Field(description="해당 설계가 요건을 충족하는 기술적 근거")

class DiagramNode(BaseModel):
    node_id: str = Field(description="영문 고유 ID (예: FW_OT, CLAROTY_CTD)")
    label: str = Field(description="도형 안에 표시될 짧은 장비/자산군 이름")
    node_type: str = Field(description="장비 유형: 'security'(보안솔루션), 'network'(방화벽/스위치/망연계), 'asset'(식별된 IP대역 자산군/서버/설비) 중 택1")
    icon_type: str = Field(default="server", description="아이콘 유형: 'firewall'(방화벽/스위치), 'diode'(일방향망연계), 'sensor'(Claroty CTD/xDome/Medigate/Edge), 'remote'(SRA/ZTNA/점프호스트), 'endpoint'(백신/매체제어/PAM), 'plc'(PLC/RTU/의료기기), 'server'(SCADA/HMI/업무서버/클라우드) 중 택1")

class DiagramZone(BaseModel):
    zone_name: str = Field(description="계층/세그먼트 구역 이름")
    nodes: List[DiagramNode] = Field(description="해당 구역 안에 배치될 장비/솔루션/식별자산군 목록 (구역당 최대 5개 권장)")

class DiagramEdge(BaseModel):
    source_id: str = Field(description="출발 노드 node_id")
    target_id: str = Field(description="도착 노드 node_id")
    label: str = Field(description="연결선 라벨")
    link_type: str = Field(default="it_1g", description="통신 유형 및 대역폭 구분: 'backbone_10g'(10G 백본/업링크), 'it_1g'(1G IT 업무/관리 통신), 'ot_control'(OT 제어/설비 프로토콜 통신), 'span_mirror'(SPAN/TAP 수동 미러링), 'unidirectional'(일방향 망연계/SRA 통제 격리) 중 택1")

class ArchitectureDesignOutput(BaseModel):
    assistant_reply: str = Field(description="사용자의 요청 또는 수정 지시에 대한 아키텍트의 답변 및 반영/변경 내역 설명")
    executive_summary: str = Field(description="고객사 환경·자산 분석 및 To-Be 세그멘테이션 설계 핵심 요약")
    identified_zones: List[str] = Field(description="보안 표준 및 IP 대역 기준으로 분리된 Zone/등급 영역 목록")
    mermaid_code: str = Field(description="렌더링 가능한 순수 Mermaid.js 코드 (```mermaid 태그 제외)")
    ppt_zones: List[DiagramZone] = Field(description="PowerPoint 슬라이드 위에서 아래 순서로 배치할 계층별 구역 및 노드 정보")
    ppt_edges: List[DiagramEdge] = Field(description="PowerPoint 도형 간 연결선 정보 (핵심 연결 8~12개 이내)")
    compliance_matrix: List[ComplianceItem] = Field(description="규제 충족 및 세그멘테이션 정책 매트릭스")

SYSTEM_PROMPT = """
당신은 15년 이상 경력의 IT/OT 융합 보안 수석 아키텍트입니다.
사용자가 좌측 패널에서 선택한 [적용 규제], [기존 보유 솔루션], [도입 예정 솔루션] 조건과 채팅창으로 전달한 [첨부 파일(RFP, As-Is 구성도, Claroty 자산 엑셀 요약)] 및 [대화형 요청/수정 지시]를 모두 종합하여 최적의 To-Be 보안 아키텍처를 설계합니다.

[설계 및 수정 원칙]
1. 선택 조건 + 채팅 지시 통합 반영:
   - 좌측 패널에서 선택된 [기존 보유 솔루션]과 [도입 예정 솔루션], [적용 규제]를 반드시 설계에 반영합니다.
   - 이전 버전 설계 JSON이 함께 전달되면, 사용자가 채팅으로 수정을 요청한 부분이나 사이드바에서 새로 변경한 솔루션/규제를 반영하여 기존 구조를 정밀하게 업데이트합니다.
2. Claroty 식별 자산 세그멘테이션:
   - 자산 요약 데이터가 있으면 IP 서브넷(/24)과 용도별로 Zone을 분리하고, 노드(`asset`)에 실제 IP 대역과 대표 자산·수량을 표기합니다.
3. OT 및 한국/국제 규제 표준 배치 규칙:
   - OT 환경: Purdue Model(Level 0~4) 기준 Level 3(OT)와 Level 4(IT) 사이에 Level 3.5(IDMZ) 배치.
   - N2SF 선택 시: 기밀(C), 민감(S), 공개(O) 등급 영역 분리 및 등급 간 데이터 통제 설계.
   - Claroty CTD/xDome/Medigate는 백본 SPAN/TAP 미러링, Claroty Edge는 SPAN 불가 구간 능동 수집, Claroty SRA는 IDMZ 배치, 일방향 망연계는 OT→IT 단방향 전송으로 배치합니다.
4. 출력 형식:
   - `mermaid_code`: `flowchart TB` 방향, 노드 ID는 영문/숫자만 사용하고 한글 라벨은 따옴표(`["한글 라벨"]`)로 감쌉니다.
   - `ppt_zones`: 위에서 아래 순서로 3~5개 구역, 구역당 2~5개 노드로 구성합니다.
"""

# ==========================================
# 4. 헬퍼 함수 (엑셀 전처리, Mermaid, PPT 생성)
# ==========================================
def summarize_claroty_assets(file_bytes: bytes, filename: str) -> str:
    try:
        buf = io.BytesIO(file_bytes)
        df = pd.read_csv(buf) if filename.lower().endswith(".csv") else pd.read_excel(buf)
        if df.empty:
            return "(빈 자산 파일)"

        cols = {str(c).lower().strip(): c for c in df.columns}
        ip_col = next((cols[k] for k in cols if "ip" in k), None)
        type_col = next((cols[k] for k in cols if any(w in k for w in ["type", "category", "유형", "분류", "role", "device"])), None)
        vendor_col = next((cols[k] for k in cols if any(w in k for w in ["vendor", "manufacturer", "제조사"])), None)
        proto_col = next((cols[k] for k in cols if any(w in k for w in ["protocol", "프로토콜"])), None)

        lines = [f"[{filename}] 총 식별 자산 수: {len(df)}대"]
        if ip_col:
            df["Subnet_24"] = df[ip_col].astype(str).str.extract(r"^(\d+\.\d+\.\d+)\.")[0].fillna("기타") + ".0/24"
            group_cols = ["Subnet_24"] + ([type_col] if type_col else []) + ([vendor_col] if vendor_col else [])
            grouped = df.groupby(group_cols, dropna=False).size().reset_index(name="수량").sort_values(by="수량", ascending=False).head(30)
            for _, row in grouped.iterrows():
                desc = " / ".join([f"{c}: {row[c]}" for c in group_cols if pd.notna(row[c])])
                lines.append(f"  * {desc} -> {row['수량']}대")
        if proto_col:
            lines.append(f"  * 주요 프로토콜: {df[proto_col].dropna().astype(str).value_counts().head(8).to_dict()}")
        return "\n".join(lines)
    except Exception as e:
        return f"[{filename}] 엑셀 파싱 참고 오류: {e}"

def render_mermaid(code: str, key_suffix: str):
    # 1. 제미나이가 실수로 넣을 수 있는 마크다운 코드블록 태그 제거
    clean_code = code.replace("```mermaid", "").replace("```", "").strip()
    
    # 2. mermaid.ink를 이용해 실제 이미지(SVG) URL 생성 (브라우저 JS 오류 원천 차단)
    encoded = base64.urlsafe_b64encode(clean_code.encode("utf-8")).decode("ascii")
    svg_url = f"https://mermaid.ink/svg/{encoded}"
    
    st.markdown(f"![To-Be 보안 아키텍처 구성도]({svg_url})")
    
    with st.expander("🔍 만약 위 그림이 깨져 보인다면? (원본 코드 확인)"):
        st.caption("문법 오류 시 아래 코드를 복사해 mermaid.live에 붙여넣으면 어느 줄이 잘못되었는지 바로 확인할 수 있습니다.")
        st.code(clean_code, language="mermaid")

def build_editable_pptx(result: ArchitectureDesignOutput) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    NAVY = RGBColor(15, 23, 42)
    ZONE_BG = RGBColor(248, 250, 252)
    ZONE_BORDER = RGBColor(203, 213, 225)
    WHITE = RGBColor(255, 255, 255)
    DARK_TEXT = RGBColor(30, 41, 59)
    MUTED_TEXT = RGBColor(100, 116, 139)

    # 슬라이드 1: 요약
    s1 = prs.slides.add_slide(blank_layout)
    h1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.85))
    h1.fill.solid(); h1.fill.fore_color.rgb = NAVY; h1.line.fill.background()
    h1.text_frame.text = "  1. IT/OT 보안 To-Be 아키텍처 및 세그멘테이션 설계 요약"
    h1.text_frame.paragraphs[0].font.size = Pt(20); h1.text_frame.paragraphs[0].font.bold = True; h1.text_frame.paragraphs[0].font.color.rgb = WHITE

    b_sum = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.15), Inches(12.1), Inches(2.6))
    b_sum.fill.solid(); b_sum.fill.fore_color.rgb = ZONE_BG; b_sum.line.color.rgb = ZONE_BORDER
    tf_s = b_sum.text_frame; tf_s.word_wrap = True
    tf_s.paragraphs[0].text = "■ 핵심 아키텍처 설계 방향"; tf_s.paragraphs[0].font.size = Pt(14); tf_s.paragraphs[0].font.bold = True; tf_s.paragraphs[0].font.color.rgb = NAVY
    p_b = tf_s.add_paragraph(); p_b.text = result.executive_summary; p_b.font.size = Pt(11.5); p_b.font.color.rgb = DARK_TEXT; p_b.space_before = Pt(8)

    b_z = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.05), Inches(12.1), Inches(2.9))
    b_z.fill.solid(); b_z.fill.fore_color.rgb = WHITE; b_z.line.color.rgb = ZONE_BORDER
    tf_z = b_z.text_frame; tf_z.word_wrap = True
    tf_z.paragraphs[0].text = "■ 식별된 보안 구역 (Zones / 등급 분리)"; tf_z.paragraphs[0].font.size = Pt(14); tf_z.paragraphs[0].font.bold = True; tf_z.paragraphs[0].font.color.rgb = NAVY
    for zt in result.identified_zones[:6]:
        pz = tf_z.add_paragraph(); pz.text = f"• {zt}"; pz.font.size = Pt(11.5); pz.font.color.rgb = DARK_TEXT; pz.space_before = Pt(4)

    # 슬라이드 2: 모던 카드형 To-Be 구성도 (개별 편집 가능)
    s2 = prs.slides.add_slide(blank_layout)
    h2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.62))
    h2.fill.solid(); h2.fill.fore_color.rgb = NAVY; h2.line.fill.background()
    h2.text_frame.text = "  2. To-Be 보안 아키텍처 구성도 (모던 카드형 · 직각 배선 · 개별 편집 가능)"
    h2.text_frame.paragraphs[0].font.size = Pt(17); h2.text_frame.paragraphs[0].font.bold = True; h2.text_frame.paragraphs[0].font.color.rgb = WHITE

    LINK_STYLES = {
        "backbone_10g":   (RGBColor(15, 23, 42),   Pt(2.25), MSO_LINE_DASH_STYLE.SOLID,     "━ 10G 백본"),
        "it_1g":          (RGBColor(37, 99, 235),  Pt(1.5),  MSO_LINE_DASH_STYLE.SOLID,     "━ 1G IT통신"),
        "ot_control":     (RGBColor(5, 150, 105),  Pt(1.75), MSO_LINE_DASH_STYLE.SOLID,     "━ OT 제어망"),
        "span_mirror":    (RGBColor(124, 58, 237), Pt(1.5),  MSO_LINE_DASH_STYLE.DASH,      "┄ SPAN 미러링"),
        "unidirectional": (RGBColor(234, 88, 12),  Pt(1.75), MSO_LINE_DASH_STYLE.ROUND_DOT, "┈ 일방향/통제"),
    }
    leg_x = 5.4
    for _, (l_rgb, _, _, l_txt) in LINK_STYLES.items():
        l_box = s2.shapes.add_textbox(Inches(leg_x), Inches(0.66), Inches(1.52), Inches(0.25))
        p_lg = l_box.text_frame.paragraphs[0]
        p_lg.text = l_txt; p_lg.font.size = Pt(8.5); p_lg.font.bold = True; p_lg.font.color.rgb = l_rgb
        leg_x += 1.54

    zones = result.ppt_zones[:5]
    num_zones = max(len(zones), 1)
    start_y, total_h, gap_y = 0.98, 6.28, 0.24
    zone_h = (total_h - gap_y * (num_zones - 1)) / num_zones
    node_anchors = {}

    # 통일된 모던 카드 스타일 (상단 컬러 포인트 태그 + 깔끔한 흰색 본체)
    CARD_ACCENTS = {
        "firewall": (RGBColor(220, 38, 38),  "FIREWALL"),
        "diode":    (RGBColor(2, 132, 199),  "DATA DIODE"),
        "sensor":   (RGBColor(5, 150, 105),  "OT VISIBILITY"),
        "remote":   (RGBColor(124, 58, 237), "SECURE REMOTE"),
        "endpoint": (RGBColor(13, 148, 136), "ENDPOINT"),
        "plc":      (RGBColor(71, 85, 105),  "OT ASSET"),
        "server":   (RGBColor(37, 99, 235),  "SYSTEM / SW"),
    }

    for idx, zone in enumerate(zones):
        zy = start_y + idx * (zone_h + gap_y)
        zb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.35), Inches(zy), Inches(12.63), Inches(zone_h))
        zb.fill.solid(); zb.fill.fore_color.rgb = ZONE_BG; zb.line.color.rgb = ZONE_BORDER; zb.line.width = Pt(1.0)

        zl = s2.shapes.add_textbox(Inches(0.48), Inches(zy + 0.08), Inches(2.15), Inches(zone_h - 0.16))
        zl.text_frame.word_wrap = True
        zl.text_frame.paragraphs[0].text = zone.zone_name; zl.text_frame.paragraphs[0].font.size = Pt(10); zl.text_frame.paragraphs[0].font.bold = True; zl.text_frame.paragraphs[0].font.color.rgb = NAVY

        nodes = zone.nodes[:5]
        n_count = max(len(nodes), 1)
        area_x, area_w = 2.75, 10.05
        # 노드 간 수평 간격을 넓게 확보해 라벨 겹침 방지
        gap_x = 0.65 if n_count <= 4 else 0.38
        node_w = min(1.92, (area_w - gap_x * (n_count - 1)) / n_count)
        node_h = min(0.72, zone_h - 0.22)
        node_y = zy + (zone_h - node_h) / 2

        for n_idx, node in enumerate(nodes):
            nx = area_x + n_idx * (node_w + gap_x)
            accent_rgb, tag_txt = CARD_ACCENTS.get(
                getattr(node, "icon_type", "server"),
                (RGBColor(71, 85, 105), "NODE")
            )

            # 1) 깔끔한 흰색 둥근 사각형 카드 본체
            card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(nx), Inches(node_y + 0.06), Inches(node_w), Inches(node_h - 0.06))
            card.fill.solid(); card.fill.fore_color.rgb = WHITE
            card.line.color.rgb = accent_rgb; card.line.width = Pt(1.5)
            tf_c = card.text_frame
            tf_c.word_wrap = True
            tf_c.margin_left = tf_c.margin_right = Inches(0.06)
            tf_c.margin_top = Inches(0.10); tf_c.margin_bottom = Inches(0.04)
            p_n = tf_c.paragraphs[0]
            p_n.text = node.label; p_n.font.size = Pt(8.2); p_n.font.bold = True; p_n.font.color.rgb = DARK_TEXT; p_n.alignment = PP_ALIGN.CENTER

            # 2) 상단 컬러 역할 배지 (알약형)
            pill = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(nx + 0.14), Inches(node_y - 0.03), Inches(node_w - 0.28), Inches(0.18))
            pill.fill.solid(); pill.fill.fore_color.rgb = accent_rgb; pill.line.fill.background()
            tf_p = pill.text_frame
            tf_p.margin_left = tf_p.margin_right = tf_p.margin_top = tf_p.margin_bottom = 0
            pp = tf_p.paragraphs[0]
            pp.text = tag_txt; pp.font.size = Pt(6.2); pp.font.bold = True; pp.font.color.rgb = WHITE; pp.alignment = PP_ALIGN.CENTER

            node_anchors[node.node_id] = {
                "zone_idx": idx,
                "cx": nx + node_w / 2,
                "cy": node_y + node_h / 2,
                "top": (nx + node_w / 2, node_y - 0.03),
                "bottom": (nx + node_w / 2, node_y + node_h),
                "left": (nx, node_y + node_h / 2 + 0.03),
                "right": (nx + node_w, node_y + node_h / 2 + 0.03),
            }

    # 연결선 및 겹침 없는 라벨 렌더링
    for e_idx, edge in enumerate(result.ppt_edges[:12]):
        if edge.source_id in node_anchors and edge.target_id in node_anchors:
            src = node_anchors[edge.source_id]
            dst = node_anchors[edge.target_id]
            l_rgb, l_width, l_dash, _ = LINK_STYLES.get(
                getattr(edge, "link_type", "it_1g"),
                LINK_STYLES["it_1g"]
            )

            is_horizontal = (src["zone_idx"] == dst["zone_idx"])
            if is_horizontal:
                if src["cx"] <= dst["cx"]:
                    x1, y1 = src["right"]; x2, y2 = dst["left"]
                else:
                    x1, y1 = src["left"]; x2, y2 = dst["right"]
                conn_type = MSO_CONNECTOR.STRAIGHT
            elif src["zone_idx"] < dst["zone_idx"]:
                x1, y1 = src["bottom"]; x2, y2 = dst["top"]
                conn_type = MSO_CONNECTOR.ELBOW
            else:
                x1, y1 = src["top"]; x2, y2 = dst["bottom"]
                conn_type = MSO_CONNECTOR.ELBOW

            conn = s2.shapes.add_connector(conn_type, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
            conn.line.color.rgb = l_rgb
            conn.line.width = l_width
            conn.line.dash_style = l_dash

            if edge.label:
                short_lbl = edge.label[:14] + (".." if len(edge.label) > 14 else "")
                if is_horizontal:
                    # 수평 연결선은 카드 사이 좁은 틈을 가리지 않도록 선 위쪽에 투명 텍스트로 배치
                    lbl_w = max(abs(x2 - x1) + 0.25, 0.85)
                    mid_x = (x1 + x2) / 2 - lbl_w / 2
                    mid_y = y1 - 0.22
                    tb = s2.shapes.add_textbox(Inches(mid_x), Inches(mid_y), Inches(lbl_w), Inches(0.20))
                    tf_e = tb.text_frame
                    tf_e.margin_left = tf_e.margin_right = tf_e.margin_top = tf_e.margin_bottom = 0
                    p_e = tf_e.paragraphs[0]
                    p_e.text = short_lbl; p_e.font.size = Pt(6.8); p_e.font.bold = True; p_e.font.color.rgb = l_rgb; p_e.alignment = PP_ALIGN.CENTER
                else:
                    # 수직(층간) 연결선은 층 사이 여백(gap_y) 중앙에 컴팩트한 캡슐 배치
                    offset_x = ((e_idx % 3) - 1) * 0.22
                    mid_x = (x1 + x2) / 2 - 0.55 + offset_x
                    mid_y = (y1 + y2) / 2 - 0.09
                    lbl_bg = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(mid_x), Inches(mid_y), Inches(1.10), Inches(0.19))
                    lbl_bg.fill.solid(); lbl_bg.fill.fore_color.rgb = WHITE
                    lbl_bg.line.color.rgb = l_rgb; lbl_bg.line.width = Pt(0.6)
                    tf_e = lbl_bg.text_frame
                    tf_e.margin_left = tf_e.margin_right = tf_e.margin_top = tf_e.margin_bottom = 0
                    p_e = tf_e.paragraphs[0]
                    p_e.text = short_lbl; p_e.font.size = Pt(6.8); p_e.font.bold = True; p_e.font.color.rgb = l_rgb; p_e.alignment = PP_ALIGN.CENTER

    # 슬라이드 3: 컴플라이언스 매트릭스 표
    s3 = prs.slides.add_slide(blank_layout)
    h3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.75))
    h3.fill.solid(); h3.fill.fore_color.rgb = NAVY; h3.line.fill.background()
    h3.text_frame.text = "  3. 규제 충족 및 Zone 세그멘테이션 정책 매트릭스"
    h3.text_frame.paragraphs[0].font.size = Pt(20); h3.text_frame.paragraphs[0].font.bold = True; h3.text_frame.paragraphs[0].font.color.rgb = WHITE

    items = result.compliance_matrix[:8]
    t_shape = s3.shapes.add_table(len(items) + 1, 5, Inches(0.4), Inches(1.0), Inches(12.53), Inches(6.0))
    table = t_shape.table
    for i, w in enumerate([Inches(1.9), Inches(1.7), Inches(2.8), Inches(3.0), Inches(3.13)]):
        table.columns[i].width = w
    for c_idx, h_text in enumerate(["조항/세그먼트 구간", "보안 영역", "As-Is 취약점/혼재", "To-Be 통제/허용정책", "충족 근거"]):
        cell = table.cell(0, c_idx); cell.fill.solid(); cell.fill.fore_color.rgb = NAVY; cell.text = h_text
        cell.text_frame.paragraphs[0].font.size = Pt(10); cell.text_frame.paragraphs[0].font.bold = True; cell.text_frame.paragraphs[0].font.color.rgb = WHITE
    for r_idx, item in enumerate(items, start=1):
        for c_idx, val in enumerate([item.requirement_id, item.category, item.as_is_gap, item.to_be_control, item.justification]):
            cell = table.cell(r_idx, c_idx); cell.fill.solid(); cell.fill.fore_color.rgb = ZONE_BG if r_idx % 2 == 0 else WHITE; cell.text = val
            cell.text_frame.paragraphs[0].font.size = Pt(9); cell.text_frame.paragraphs[0].font.color.rgb = DARK_TEXT

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue()

# ==========================================
# 5. 좌측 사이드바: 프로젝트 히스토리 불러오기 + 선택형 컨트롤
# ==========================================
SOLUTIONS_LIST = [
    "없음 (신규 구축)",
    "Claroty CTD (OT 수동형 자산 가시성/위협 탐지)",
    "Claroty xDome (클라우드 기반 OT/IoT 보안 플랫폼)",
    "Claroty Edge (SPAN 불가 구간 능동형 자산 수집)",
    "Claroty Medigate / xDome Healthcare (의료기기 IoMT 자산·위협 관리)",
    "Claroty SRA / xSA (OT 전용 안전한 원격접속)",
    "산업용 방화벽 (FortiGate Rugged / Hirschmann / Stratix)",
    "일방향 망연계 (Data Diode - 제어망→업무망 단방향 전송)",
    "양방향 망연계 (IDMZ 중계 및 패치/파일 전송)",
    "산업용 L2/L3 보안 스위치 (Cisco IE / Hirschmann / Rockwell Stratix)",
    "발전 제어망 전용 백신/화이트리스트 (Trellix / AhnLab EPS 등)",
    "매체제어 및 휴대용 저장매체(USB) 반입통제 솔루션",
    "발전소 엔지니어링 워크스테이션(EWS) 접근통제/PAM",
    "의료기기 전용 마이크로 세그멘테이션 (NAC/방화벽 정책 연동)",
    "PACS / EMR 연동 구간 전용 보안 게이트웨이",
    "일반 IT/경계 방화벽 (Palo Alto / Fortinet / SECUI 등)",
    "제로트러스트 네트워크 접근 (ZTNA / SDP)",
    "IDAM / 통합계정권한관리 (IAM / PAM)",
    "DLP / 문서중앙화 / DRM (데이터 등급별 유출방지)",
    "망분리 VDI / 논리적·물리적 망분리 게이트웨이",
    "WAF / 웹방화벽 및 DDoS 대응장비",
    "클라우드 보안 (CSPM / CWPP / 클라우드 방화벽)",
    "NAC (네트워크 접근제어)",
    "통합로그 / SIEM / SOAR 관제 플랫폼"
]

with st.sidebar:
    st.header("📂 프로젝트 및 히스토리")
    project_title = st.text_input("현재 프로젝트/고객사명", value="신규 OT 보안 설계 프로젝트")

    saved_records = load_history_records()
    if saved_records:
        options_map = {f"[{r.get('Timestamp', '')[:16]}] {r.get('Project_Title', '무제')}": r for r in saved_records}
        hist_labels = ["(새 프로젝트 작성)"] + list(options_map.keys())
        
        current_label = st.session_state.get("current_loaded_hist", "(새 프로젝트 작성)")
        default_idx = hist_labels.index(current_label) if current_label in hist_labels else 0

        selected_hist_label = st.selectbox("📂 저장된 히스토리 바로보기", hist_labels, index=default_idx)

        if selected_hist_label != current_label:
            st.session_state["current_loaded_hist"] = selected_hist_label
            if selected_hist_label == "(새 프로젝트 작성)":
                st.session_state["chat_history"] = []
                st.session_state["latest_design_json"] = None
            else:
                rec = options_map[selected_hist_label]
                st.session_state["chat_history"] = json.loads(rec.get("Chat_History_JSON", "[]"))
                st.session_state["latest_design_json"] = rec.get("Latest_Design_JSON", "")
            st.rerun()
    else:
        st.caption("💡 Google Sheets 연동 완료 시 작업 내역이 자동 저장됩니다.")

    st.markdown("---")
    st.header("1. AI 추론 모델")
    MODEL_OPTIONS = {
        "Gemini 4 Hydrogen (멀티모달 문서·도면 분석)": "gemini-4-hydrogen",
    }
    selected_model_label = st.selectbox("모델 선택", list(MODEL_OPTIONS.keys()), index=0)
    selected_model_id = MODEL_OPTIONS[selected_model_label]

    st.header("2. 보안 솔루션 선택 (자동 참조)")
    current_solutions = st.multiselect(
        "현재 보유 솔루션 (As-Is)",
        SOLUTIONS_LIST,
        default=["일반 IT/경계 방화벽 (Palo Alto / Fortinet / SECUI 등)"]
    )
    custom_current = st.text_input("기타 보유 솔루션 직접 입력", placeholder="예: 시스코 ISE, 노후 방화벽")

    planned_solutions = st.multiselect(
        "도입 예정 솔루션 (To-Be)",
        [s for s in SOLUTIONS_LIST if s != "없음 (신규 구축)"],
        default=[
            "산업용 방화벽 (FortiGate Rugged / Hirschmann / Stratix)",
            "Claroty CTD (OT 수동형 자산 가시성/위협 탐지)",
            "Claroty SRA / xSA (OT 전용 안전한 원격접속)"
        ]
    )
    custom_planned = st.text_input("기타 도입 솔루션 직접 입력", placeholder="예: 자사 특화 관제, 신규 IDS")

    st.header("3. 적용 규제 및 표준 선택")
    selected_regs = st.multiselect(
        "보안 표준 / 규제 (IT · OT · 한국 특화)",
        [
            "IEC 62443-3-3 (SL-2 / 산업제어시스템 보안)",
            "IEC 62443-3-3 (SL-3 / 고위험 기반시설)",
            "N2SF (국가 망 보안체계 - C/S/O 등급 분리)",
            "ISO/IEC 27001 (정보보호 경영시스템)",
            "ISMS-P (정보보호 및 개인정보보호 관리체계)",
            "국내 주요정보통신기반시설 보호지침",
            "제로트러스트 가이드라인 2.0 (과기정통부/KISA)",
            "전자금융감독규정 / 금융보안원 가이드라인",
            "NIS2 Directive (EU 네트워크 및 정보보안 지침)",
            "IAEA NSS 17-T (원자력/발전 설비 사이버보안)",
            "의료법 / 의료기기 사이버보안 허가·심사 가이드라인"
        ],
        default=["IEC 62443-3-3 (SL-2 / 산업제어시스템 보안)", "국내 주요정보통신기반시설 보호지침"]
    )

    st.markdown("---")
    if st.button("🗑️ 현재 화면 대화 초기화", use_container_width=True):
        st.session_state["chat_history"] = []
        st.session_state["latest_design_json"] = None
        st.rerun()

# ==========================================
# 6. 메인 채팅 워크스페이스
# ==========================================
st.title("🛡️ IT/OT 보안 To-Be 아키텍트 AI (하이브리드)")
st.caption("좌측에서 **프로젝트명·규제·솔루션**을 설정하고, 하단 채팅창의 **클립(📎)**으로 파일을 첨부하면 자동 설계 및 Google Sheets 히스토리 저장이 진행됩니다.")

# 세션 초기화 시 Google Sheets의 가장 최근 히스토리 자동 복원
if "chat_history" not in st.session_state:
    recent_records = load_history_records()
    if recent_records:
        try:
            latest_rec = recent_records[0]
            st.session_state["chat_history"] = json.loads(latest_rec.get("Chat_History_JSON", "[]"))
            st.session_state["latest_design_json"] = latest_rec.get("Latest_Design_JSON", "")
            st.session_state["current_loaded_hist"] = f"[{latest_rec.get('Timestamp', '')[:16]}] {latest_rec.get('Project_Title', '무제')}"
        except Exception:
            st.session_state["chat_history"] = []
            st.session_state["latest_design_json"] = None
    else:
        st.session_state["chat_history"] = []
        st.session_state["latest_design_json"] = None

if "latest_design_json" not in st.session_state:
    st.session_state["latest_design_json"] = None

def render_design_card(result_dict: dict, msg_idx: int):
    result_obj = ArchitectureDesignOutput(**result_dict)
    st.markdown(result_obj.assistant_reply)

    pptx_bytes = build_editable_pptx(result_obj)
    st.download_button(
        label=f"📊 편집 가능한 PowerPoint(.pptx) 다운로드 (버전 #{msg_idx})",
        data=pptx_bytes,
        file_name=f"{project_title}_ToBe_v{msg_idx}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key=f"dl_pptx_{msg_idx}",
        type="primary"
    )

    t1, t2, t3 = st.tabs(["📊 구성도 미리보기", "📋 규제·세그멘테이션 매트릭스", "📝 설계 요약 및 Mermaid 코드"])
    with t1:
        render_mermaid(result_obj.mermaid_code, key_suffix=str(msg_idx))
    with t2:
        df = pd.DataFrame([item.model_dump() for item in result_obj.compliance_matrix])
        df.columns = ["조항/세그먼트 구간", "보안 영역", "As-Is 취약점/혼재", "To-Be 통제/허용정책", "충족 근거"]
        st.dataframe(df, use_container_width=True)
    with t3:
        st.write(result_obj.executive_summary)
        for z in result_obj.identified_zones:
            st.markdown(f"- {z}")
        st.code(result_obj.mermaid_code, language="mermaid")

for idx, msg in enumerate(st.session_state["chat_history"]):
    with st.chat_message(msg["role"]):
        if msg["role"] == "user":
            if msg.get("files"):
                st.caption("📎 첨부 파일: " + ", ".join(msg["files"]))
            st.markdown(msg["text"])
        else:
            render_design_card(msg["design_data"], msg_idx=idx)

user_msg = st.chat_input(
    "설계 요구사항을 입력하거나, 생성된 구성도의 수정 지시를 입력하세요 (좌측 선택 조건 자동 참조)...",
    accept_file="multiple",
    file_type=["pdf", "png", "jpg", "jpeg", "xlsx", "csv"]
)

if user_msg:
    text_input = user_msg.text.strip() if user_msg.text else ""
    uploaded_files = user_msg.files if user_msg.files else []

    if not text_input and not uploaded_files:
        st.stop()

    all_current = list(current_solutions)
    if custom_current.strip():
        all_current.extend([i.strip() for i in custom_current.split(",") if i.strip()])

    all_planned = list(planned_solutions)
    if custom_planned.strip():
        all_planned.extend([i.strip() for i in custom_planned.split(",") if i.strip()])

    file_names = [f.name for f in uploaded_files]
    display_text = text_input if text_input else "좌측에 선택한 솔루션·규제 조건과 첨부 파일을 바탕으로 To-Be 아키텍처를 설계해 주세요."

    with st.chat_message("user"):
        if file_names:
            st.caption("📎 첨부 파일: " + ", ".join(file_names))
        st.markdown(display_text)

    st.session_state["chat_history"].append({
        "role": "user",
        "text": display_text,
        "files": file_names
    })

    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        st.error("GEMINI_API_KEY가 설정되지 않았습니다. Streamlit Secrets 설정을 확인해 주세요.")
        st.stop()

    client = genai.Client(api_key=api_key)
    gemini_files = []
    temp_paths = []
    excel_summaries = []

    with st.chat_message("assistant"):
        try:
            with st.spinner("사이드바 선택 조건 및 첨부 파일 분석 중..."):
                for uf in uploaded_files:
                    fname = uf.name
                    fbytes = uf.getvalue()
                    if fname.lower().endswith((".xlsx", ".csv")):
                        excel_summaries.append(summarize_claroty_assets(fbytes, fname))
                    else:
                        suffix = os.path.splitext(fname)[1]
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                            tmp.write(fbytes)
                            temp_paths.append(tmp.name)
                        gemini_files.append(client.files.upload(file=temp_paths[-1]))

                prev_context = ""
                if st.session_state["latest_design_json"]:
                    prev_context = f"\n[이전 버전 To-Be 설계 JSON (수정 요청 또는 변경된 사이드바 조건을 이 구조에 반영할 것)]:\n{st.session_state['latest_design_json']}\n"

                excel_context = ""
                if excel_summaries:
                    excel_context = "\n[첨부된 Claroty 식별 자산 인벤토리 요약]:\n" + "\n\n".join(excel_summaries) + "\n"

                full_prompt = f"""
                [프로젝트명]: {project_title}
                [좌측 사이드바 선택 조건 (필수 반영)]
                - 적용 규제 및 표준: {', '.join(selected_regs) if selected_regs else 'IEC 62443 기본 기준'}
                - 현재 보유 솔루션 (As-Is): {', '.join(all_current) if all_current else '없음'}
                - 도입 예정 솔루션 (To-Be): {', '.join(all_planned) if all_planned else '규제 충족에 필요한 표준 솔루션 자동 추천'}
                {prev_context}
                {excel_context}
                [사용자 채팅 지시사항]
                - {display_text}
                """

                # 503 트래픽 폭주 대비 점진적 재시도(Backoff) 및 안정화 모델 우회 로직
                all_candidates = [
                    "gemini-4-hydrogen",
                    "gemini-4-hydrogen",
                    "gemini-2.5-flash",
                ]
                fallback_models = [selected_model_id] + [m for m in all_candidates if m != selected_model_id]
                # 동일 모델 연속 재시도 보장
                if selected_model_id == "gemini-4-hydrogen":
                    fallback_models = ["gemini-4-hydrogen", "gemini-4-hydrogen", "gemini-2.5-flash"]

                response = None
                last_err = None

                for idx_m, m_id in enumerate(fallback_models):
                    try:
                        if idx_m > 0:
                            st.info(f"⏳ 서버 혼잡으로 재시도/보조 모델(`{m_id}`)로 자동 전환하여 진행합니다... ({idx_m}회차)")
                        response = client.models.generate_content(
                            model=m_id,
                            contents=gemini_files + [full_prompt],
                            config=types.GenerateContentConfig(
                                system_instruction=SYSTEM_PROMPT,
                                temperature=0.2,
                                response_mime_type="application/json",
                                response_schema=ArchitectureDesignOutput,
                            ),
                        )
                        break
                    except Exception as api_err:
                        last_err = api_err
                        err_msg = str(api_err)
                        if any(code in err_msg for code in ["503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"]):
                            time.sleep(2 * (idx_m + 1))  # 2초 -> 4초 -> 6초 점진적 대기
                            continue
                        raise api_err

                if response is None:
                    raise last_err

                result_obj: ArchitectureDesignOutput = response.parsed
                result_dict = result_obj.model_dump()
                design_json_str = json.dumps(result_dict, ensure_ascii=False)

                st.session_state["latest_design_json"] = design_json_str
                new_idx = len(st.session_state["chat_history"])
                st.session_state["chat_history"].append({
                    "role": "assistant",
                    "design_data": result_dict
                })

                # Google Sheets에 자동 저장
                saved_ok = save_history_to_sheet(project_title, display_text, st.session_state["chat_history"], design_json_str)
                if saved_ok:
                    st.toast("✅ Google Sheets에 설계 히스토리가 자동 저장되었습니다.")

                render_design_card(result_dict, msg_idx=new_idx)

        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")

        finally:
            for gf in gemini_files:
                try:
                    client.files.delete(name=gf.name)
                except Exception:
                    pass
            for tp in temp_paths:
                if os.path.exists(tp):
                    os.remove(tp)
