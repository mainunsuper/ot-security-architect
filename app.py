import io
import os
import tempfile
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from typing import List
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# 페이지 기본 설정
st.set_page_config(page_title="IT/OT 보안 To-Be 아키텍트 AI", layout="wide")

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
# 2. 메인 화면 및 데이터 규격 정의
# ==========================================
st.title("🛡️ IT/OT 보안 To-Be 아키텍처 자동 설계 에이전트")
st.markdown("RFP 문서, 기존 구성도, **Claroty 식별 자산 엑셀**을 분석하여 **세그멘테이션된 PowerPoint(.pptx) 구성도**와 **컴플라이언스 매트릭스**를 자동 생성합니다.")

class ComplianceItem(BaseModel):
    requirement_id: str = Field(description="RFP 요구사항 번호, 표준 조항 또는 세그멘테이션 구간 (예: IEC 62443-3-3 SR 5.1, Zone1→Zone2 정책)")
    category: str = Field(description="보안 영역 (예: 마이크로 세그멘테이션, 네트워크 분리, 자산 가시성, 원격접속 통제)")
    as_is_gap: str = Field(description="기존 As-Is 구성 또는 식별 자산의 취약점/혼재 현황")
    to_be_control: str = Field(description="To-Be 구성도에 반영된 보안 통제 방안, 솔루션 및 허용 프로토콜")
    justification: str = Field(description="해당 설계가 요건을 충족하는 기술적 근거")

class DiagramNode(BaseModel):
    node_id: str = Field(description="영문 고유 ID (예: FW_OT, CLAROTY_CTD, ZONE_PLC_20)")
    label: str = Field(description="도형 안에 표시될 짧은 장비/자산군 이름 (예: 10.10.20.0/24 PLC 18대, Claroty CTD)")
    node_type: str = Field(description="장비 유형: 'security'(보안솔루션), 'network'(방화벽/스위치/망연계), 'asset'(식별된 IP대역 자산군/서버/설비) 중 택1")

class DiagramZone(BaseModel):
    zone_name: str = Field(description="계층/세그먼트 구역 이름 (예: Level 3.5 - IDMZ, Zone 1 - 터빈 제어 PLC망)")
    nodes: List[DiagramNode] = Field(description="해당 구역 안에 배치될 장비/솔루션/식별자산군 목록 (구역당 최대 5개 권장)")

class DiagramEdge(BaseModel):
    source_id: str = Field(description="출발 노드 node_id")
    target_id: str = Field(description="도착 노드 node_id")
    label: str = Field(description="연결선 라벨 (예: SPAN 미러링, S7comm 허용, 일방향 전송)")

class ArchitectureDesignOutput(BaseModel):
    executive_summary: str = Field(description="고객사 환경·자산 분석 및 To-Be 세그멘테이션 설계 핵심 요약")
    identified_zones: List[str] = Field(description="보안 표준 및 IP 대역 기준으로 분리된 Zone/등급 영역 목록")
    mermaid_code: str = Field(description="렌더링 가능한 순수 Mermaid.js 코드 (```mermaid 태그 제외)")
    ppt_zones: List[DiagramZone] = Field(description="PowerPoint 슬라이드 위에서 아래(Top-to-Bottom) 순서로 배치할 계층별 구역 및 노드 정보")
    ppt_edges: List[DiagramEdge] = Field(description="PowerPoint 도형 간 연결선 정보 (핵심 연결 8~12개 이내)")
    compliance_matrix: List[ComplianceItem] = Field(description="규제 충족 및 세그멘테이션 정책 매트릭스")

SYSTEM_PROMPT = """
당신은 15년 이상 경력의 IT/OT 융합 보안 수석 아키텍트입니다.
사용자가 제공하는 고객사 RFP(문서), 기존 As-Is 네트워크 구성도, Claroty 식별 자산 인벤토리(IP 대역 및 용도별 요약), 적용 규제 기준, 그리고 [기존/도입 예정 솔루션] 정보를 종합하여 최적의 To-Be 보안 아키텍처를 설계합니다.

[설계 및 세그멘테이션 원칙]
1. Claroty 식별 자산 자동 세그멘테이션:
   - 자산 인벤토리 데이터가 제공되면, 식별된 IP 서브넷(/24)과 장비 용도(예: PLC/RTU 제어망, HMI/SCADA 운영망, 의료기기 IoMT망, 일반 PC/서버망 등)를 분석해 위험도와 기능별로 Zone을 명확히 분리(Segmentation)합니다.
   - 각 Zone 내부 노드(`asset`)에는 실제 식별된 IP 대역과 대표 자산명·수량을 표기합니다 (예: `10.10.20.0/24 [Siemens PLC 18대]`).
   - 서로 다른 Zone 간 경계에는 산업용 방화벽 또는 L3 보안 스위치를 배치하고, 컴플라이언스 매트릭스에 Zone 간 최소 권한 통신(허용 프로토콜) 권고안을 포함합니다.
2. OT/제어시스템 계층 구조: Purdue Model(Level 0~4)을 기반으로 Level 3(OT)와 Level 4(IT) 사이에 반드시 Level 3.5(IDMZ)를 배치합니다.
3. 한국 특화 및 국제 표준 적용:
   - 'N2SF (국가 망 보안체계)' 선택 시: 기밀(C), 민감(S), 공개(O) 등급 영역으로 분리하고 등급 간 데이터 통제를 설계합니다.
   - 'ISO/IEC 27001' 또는 'ISMS-P' 선택 시: 접근통제(IAM/PAM), 암호화, 로그 모니터링(SIEM/SOAR), 네트워크 분리 통제를 반영합니다.
4. [세부 솔루션별 표준 배치 규칙]
   - Claroty CTD / xDome / Medigate: 코어 및 주요 L2/L3 백본 스위치에서 SPAN/TAP 미러링으로 트래픽을 수동 수집하도록 배치합니다.
   - Claroty Edge: SPAN 포트 구성이 어려운 원격지나 독립 의료기기 네트워크에 배치하여 능동 수집 후 중앙 서버로 전송하게 그립니다.
   - Claroty SRA / xSA: 반드시 Level 3.5(IDMZ)에 배치하여 외부 벤더/유지보수 인력이 경유하도록 구성합니다.
   - 일방향 망연계(Data Diode): Level 3(OT)에서 Level 4(IT)로 데이터를 보낼 때 단방향으로 배치합니다.

[출력 작성 규칙]
- `mermaid_code`: `flowchart TB` 방향, 노드 ID는 영문/숫자만 사용하고 한글 라벨은 따옴표(`["한글 라벨"]`)로 감쌉니다.
- `ppt_zones`: 파워포인트 한 장에 깔끔하게 들어가도록 위(상위망/외부망/Level 4)에서 아래(현장망/세그먼트 Zone) 순서로 3~5개 구역으로 정리하고, 각 구역 내 노드는 2~5개로 압축해 가독성을 높입니다.
"""

# ==========================================
# Claroty 자산 인벤토리 엑셀/CSV 전처리 함수
# ==========================================
def summarize_claroty_assets(uploaded_excel) -> str:
    """Claroty 자산 엑셀을 읽어 IP 서브넷(/24) 및 장비 유형·용도별로 그룹핑해 요약 텍스트로 반환합니다."""
    if uploaded_excel is None:
        return ""
    try:
        if uploaded_excel.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_excel)
        else:
            df = pd.read_excel(uploaded_excel)

        if df.empty:
            return "(업로드된 자산 파일에 데이터가 없습니다.)"

        cols = {str(c).lower().strip(): c for c in df.columns}
        ip_col = next((cols[k] for k in cols if "ip" in k), None)
        type_col = next((cols[k] for k in cols if any(w in k for w in ["type", "category", "유형", "분류", "role", "device"])), None)
        vendor_col = next((cols[k] for k in cols if any(w in k for w in ["vendor", "manufacturer", "제조사"])), None)
        proto_col = next((cols[k] for k in cols if any(w in k for w in ["protocol", "프로토콜"])), None)
        vlan_col = next((cols[k] for k in cols if "vlan" in k), None)

        summary_lines = [f"- 총 식별 자산 수: {len(df)}대"]

        if ip_col:
            df["Subnet_24"] = df[ip_col].astype(str).str.extract(r"^(\d+\.\d+\.\d+)\.")[0]
            df["Subnet_24"] = df["Subnet_24"].fillna("기타/비IPv4") + ".0/24"
            df["Subnet_24"] = df["Subnet_24"].str.replace("기타/비IPv4.0/24", "기타/미할당 IP", regex=False)

            group_cols = ["Subnet_24"]
            if vlan_col:
                group_cols.append(vlan_col)
            if type_col:
                group_cols.append(type_col)
            if vendor_col:
                group_cols.append(vendor_col)

            grouped = df.groupby(group_cols, dropna=False).size().reset_index(name="수량")
            grouped = grouped.sort_values(by="수량", ascending=False).head(30)

            summary_lines.append("- [IP 대역(Subnet) 및 용도별 주요 자산 그룹 분포 (상위 30개 그룹)]")
            for _, row in grouped.iterrows():
                desc = " / ".join([f"{col}: {row[col]}" for col in group_cols if pd.notna(row[col])])
                summary_lines.append(f"  * {desc} -> {row['수량']}대")
        elif type_col:
            grouped = df.groupby([type_col], dropna=False).size().reset_index(name="수량").sort_values(by="수량", ascending=False).head(20)
            summary_lines.append("- [장비 유형별 분포]")
            for _, row in grouped.iterrows():
                summary_lines.append(f"  * {row[type_col]} -> {row['수량']}대")

        if proto_col:
            top_protos = df[proto_col].dropna().astype(str).value_counts().head(10).to_dict()
            summary_lines.append(f"- [주요 식별 통신 프로토콜]: {top_protos}")

        return "\n".join(summary_lines)
    except Exception as e:
        return f"(자산 엑셀 분석 중 오류 발생: {e})"

def render_mermaid(code: str):
    html_code = f"""
    <div class="mermaid" style="background-color: white; padding: 20px; border-radius: 10px;">
        {code}
    </div>
    <script type="module">
        import mermaid from '[https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs](https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs)';
        mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
    </script>
    """
    components.html(html_code, height=600, scrolling=True)

# ==========================================
# 편집 가능한 PowerPoint(.pptx) 생성 함수
# ==========================================
def build_editable_pptx(result: ArchitectureDesignOutput) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    NAVY = RGBColor(20, 42, 76)
    BLUE_ACCENT = RGBColor(0, 102, 204)
    TEAL_SECURITY = RGBColor(13, 128, 101)
    ORANGE_NET = RGBColor(214, 112, 24)
    SLATE_ASSET = RGBColor(71, 85, 105)
    ZONE_BG = RGBColor(241, 245, 249)
    ZONE_BORDER = RGBColor(148, 163, 184)
    WHITE = RGBColor(255, 255, 255)
    DARK_TEXT = RGBColor(30, 41, 59)

    # 슬라이드 1: 설계 요약 및 보안 구역 개요
    s1 = prs.slides.add_slide(blank_layout)
    header1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.9))
    header1.fill.solid()
    header1.fill.fore_color.rgb = NAVY
    header1.line.fill.background()
    tf_h1 = header1.text_frame
    tf_h1.text = "  1. IT/OT 보안 To-Be 아키텍처 및 세그멘테이션 설계 요약"
    tf_h1.paragraphs[0].font.size = Pt(22)
    tf_h1.paragraphs[0].font.bold = True
    tf_h1.paragraphs[0].font.color.rgb = WHITE

    box_sum = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.2), Inches(12.1), Inches(2.6))
    box_sum.fill.solid()
    box_sum.fill.fore_color.rgb = ZONE_BG
    box_sum.line.color.rgb = ZONE_BORDER
    tf_sum = box_sum.text_frame
    tf_sum.word_wrap = True
    p_title = tf_sum.paragraphs[0]
    p_title.text = "■ 핵심 아키텍처 및 자산 세그멘테이션 방향"
    p_title.font.size = Pt(15)
    p_title.font.bold = True
    p_title.font.color.rgb = NAVY

    p_body = tf_sum.add_paragraph()
    p_body.text = result.executive_summary
    p_body.font.size = Pt(12)
    p_body.font.color.rgb = DARK_TEXT
    p_body.space_before = Pt(8)

    box_zone = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.1), Inches(12.1), Inches(2.9))
    box_zone.fill.solid()
    box_zone.fill.fore_color.rgb = WHITE
    box_zone.line.color.rgb = ZONE_BORDER
    tf_z = box_zone.text_frame
    tf_z.word_wrap = True
    pz_title = tf_z.paragraphs[0]
    pz_title.text = "■ 식별된 보안 구역 (Zones / IP 대역 및 등급 분리 체계)"
    pz_title.font.size = Pt(15)
    pz_title.font.bold = True
    pz_title.font.color.rgb = NAVY

    for z_text in result.identified_zones[:6]:
        pz = tf_z.add_paragraph()
        pz.text = f"• {z_text}"
        pz.font.size = Pt(12)
        pz.font.color.rgb = DARK_TEXT
        pz.space_before = Pt(4)

    # 슬라이드 2: 편집 가능한 To-Be 아키텍처 구성도
    s2 = prs.slides.add_slide(blank_layout)
    header2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.75))
    header2.fill.solid()
    header2.fill.fore_color.rgb = NAVY
    header2.line.fill.background()
    tf_h2 = header2.text_frame
    tf_h2.text = "  2. To-Be 보안 아키텍처 구성도 (모든 도형/선 개별 편집 가능)"
    tf_h2.paragraphs[0].font.size = Pt(20)
    tf_h2.paragraphs[0].font.bold = True
    tf_h2.paragraphs[0].font.color.rgb = WHITE

    legend_items = [
        ("보안솔루션", TEAL_SECURITY, Inches(8.1)),
        ("방화벽/망연계/네트워크", ORANGE_NET, Inches(9.8)),
        ("식별 IP자산군/설비", SLATE_ASSET, Inches(11.6)),
    ]
    for l_text, l_color, l_left in legend_items:
        lg = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l_left, Inches(0.85), Inches(1.55), Inches(0.32))
        lg.fill.solid()
        lg.fill.fore_color.rgb = l_color
        lg.line.fill.background()
        tf_lg = lg.text_frame
        tf_lg.text = l_text
        tf_lg.paragraphs[0].font.size = Pt(9)
        tf_lg.paragraphs[0].font.bold = True
        tf_lg.paragraphs[0].font.color.rgb = WHITE
        tf_lg.paragraphs[0].alignment = PP_ALIGN.CENTER

    zones = result.ppt_zones[:5]
    num_zones = max(len(zones), 1)
    start_y = 1.25
    total_h = 5.95
    gap_y = 0.14
    zone_h = (total_h - gap_y * (num_zones - 1)) / num_zones

    node_centers = {}

    for idx, zone in enumerate(zones):
        zy = start_y + idx * (zone_h + gap_y)
        z_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(zy), Inches(12.53), Inches(zone_h))
        z_box.fill.solid()
        z_box.fill.fore_color.rgb = ZONE_BG
        z_box.line.color.rgb = ZONE_BORDER
        z_box.line.width = Pt(1.2)

        z_label = s2.shapes.add_textbox(Inches(0.5), Inches(zy + 0.05), Inches(2.3), Inches(zone_h - 0.1))
        tf_zl = z_label.text_frame
        tf_zl.word_wrap = True
        p_zl = tf_zl.paragraphs[0]
        p_zl.text = zone.zone_name
        p_zl.font.size = Pt(11)
        p_zl.font.bold = True
        p_zl.font.color.rgb = NAVY

        nodes = zone.nodes[:5]
        n_count = max(len(nodes), 1)
        area_x = 2.9
        area_w = 9.8
        gap_x = 0.25
        node_w = min(2.1, (area_w - gap_x * (n_count - 1)) / n_count)
        node_h = min(0.72, zone_h - 0.20)
        node_y = zy + (zone_h - node_h) / 2

        for n_idx, node in enumerate(nodes):
            nx = area_x + n_idx * (node_w + gap_x)
            n_shape = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(nx), Inches(node_y), Inches(node_w), Inches(node_h))
            n_shape.fill.solid()
            if node.node_type == "security":
                n_shape.fill.fore_color.rgb = TEAL_SECURITY
            elif node.node_type == "network":
                n_shape.fill.fore_color.rgb = ORANGE_NET
            else:
                n_shape.fill.fore_color.rgb = SLATE_ASSET
            n_shape.line.color.rgb = WHITE
            n_shape.line.width = Pt(1)

            tf_n = n_shape.text_frame
            tf_n.word_wrap = True
            p_n = tf_n.paragraphs[0]
            p_n.text = node.label
            p_n.font.size = Pt(9.5)
            p_n.font.bold = True
            p_n.font.color.rgb = WHITE
            p_n.alignment = PP_ALIGN.CENTER

            node_centers[node.node_id] = (nx + node_w / 2, node_y + node_h / 2)

    for edge in result.ppt_edges[:12]:
        if edge.source_id in node_centers and edge.target_id in node_centers:
            x1, y1 = node_centers[edge.source_id]
            x2, y2 = node_centers[edge.target_id]
            conn = s2.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
            conn.line.color.rgb = BLUE_ACCENT
            conn.line.width = Pt(1.5)

            if edge.label:
                mid_x = (x1 + x2) / 2 - 0.6
                mid_y = (y1 + y2) / 2 - 0.12
                lbl = s2.shapes.add_textbox(Inches(mid_x), Inches(mid_y), Inches(1.2), Inches(0.24))
                tf_e = lbl.text_frame
                p_e = tf_e.paragraphs[0]
                p_e.text = edge.label
                p_e.font.size = Pt(8)
                p_e.font.bold = True
                p_e.font.color.rgb = BLUE_ACCENT
                p_e.alignment = PP_ALIGN.CENTER

    # 슬라이드 3: 컴플라이언스 및 세그멘테이션 정책 매트릭스 표
    s3 = prs.slides.add_slide(blank_layout)
    header3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.75))
    header3.fill.solid()
    header3.fill.fore_color.rgb = NAVY
    header3.line.fill.background()
    tf_h3 = header3.text_frame
    tf_h3.text = "  3. 규제 충족 및 Zone 세그멘테이션 정책 매트릭스"
    tf_h3.paragraphs[0].font.size = Pt(20)
    tf_h3.paragraphs[0].font.bold = True
    tf_h3.paragraphs[0].font.color.rgb = WHITE

    items = result.compliance_matrix[:8]
    rows = len(items) + 1
    cols = 5
    table_shape = s3.shapes.add_table(rows, cols, Inches(0.4), Inches(1.0), Inches(12.53), Inches(6.0))
    table = table_shape.table

    col_widths = [Inches(1.9), Inches(1.7), Inches(2.8), Inches(3.0), Inches(3.13)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = w

    headers = ["조항/세그먼트 구간", "보안 영역", "As-Is 취약점/혼재", "To-Be 통제/허용정책", "충족 근거"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        cell.text = h_text
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = WHITE

    for r_idx, item in enumerate(items, start=1):
        row_vals = [item.requirement_id, item.category, item.as_is_gap, item.to_be_control, item.justification]
        for c_idx, val in enumerate(row_vals):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ZONE_BG if r_idx % 2 == 0 else WHITE
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9)
            p.font.color.rgb = DARK_TEXT

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()

# ==========================================
# 3. 좌측 사이드바: 자료, 자산 엑셀, 규제, 솔루션 선택 UI
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
    st.header("1. 분석 자료 및 Claroty 자산 업로드")
    rfp_file = st.file_uploader("고객사 RFP 또는 요구사항 문서 (PDF)", type=["pdf"])
    asis_file = st.file_uploader("기존 As-Is 구성도 (PNG, JPG, PDF)", type=["png", "jpg", "jpeg", "pdf"])
    asset_excel = st.file_uploader(
        "Claroty 식별 자산 인벤토리 (XLSX, CSV - 선택)",
        type=["xlsx", "csv"],
        help="Claroty CTD / xDome / Medigate에서 내보낸 자산 목록을 올리면 IP 대역과 용도별로 자동 세그멘테이션하여 구성도에 배치합니다."
    )
    
    st.header("2. 보안 솔루션 구성 선택")
    current_solutions = st.multiselect(
        "현재 구성되어 있는 보안 솔루션 (As-Is)",
        SOLUTIONS_LIST,
        default=["일반 IT/경계 방화벽 (Palo Alto / Fortinet / SECUI 등)"]
    )
    custom_current = st.text_input("기타 기존 보유 솔루션 직접 입력 (선택)", placeholder="예: 시스코 ISE, 노후 방화벽 등")

    st.markdown("---")

    planned_solutions = st.multiselect(
        "도입 예정인 보안 솔루션 (To-Be)",
        [s for s in SOLUTIONS_LIST if s != "없음 (신규 구축)"],
        default=[
            "산업용 방화벽 (FortiGate Rugged / Hirschmann / Stratix)",
            "Claroty CTD (OT 수동형 자산 가시성/위협 탐지)",
            "Claroty SRA / xSA (OT 전용 안전한 원격접속)"
        ]
    )
    custom_planned = st.text_input("기타 도입 예정 솔루션 직접 입력 (선택)", placeholder="예: 자사 특화 관제, 신규 IDS 등")

    st.header("3. 규제 및 추가 요건")
    regulations = st.multiselect(
        "적용할 보안 표준 및 규제 선택 (IT / OT / 한국 특화)",
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
    extra_notes = st.text_area("추가 반영 사항 (선택)", placeholder="예: 10.10.20.0/24 대역 PLC 망과 HMI 망 분리, SPAN 불가 구간 Claroty Edge 적용 등")
    run_btn = st.button("🚀 To-Be 아키텍처 및 PPT 생성하기", type="primary", use_container_width=True)

# ==========================================
# 4. 메인 화면 실행 로직
# ==========================================
if run_btn:
    all_current = list(current_solutions)
    if custom_current.strip():
        all_current.extend([item.strip() for item in custom_current.split(",") if item.strip()])

    all_planned = list(planned_solutions)
    if custom_planned.strip():
        all_planned.extend([item.strip() for item in custom_planned.split(",") if item.strip()])

    if not rfp_file and not asis_file and not asset_excel and not extra_notes and not all_planned:
        st.warning("분석할 파일(RFP, 구성도, 자산 엑셀)이나 솔루션, 추가 요구사항 중 최소 1가지를 입력해 주세요.")
        st.stop()

    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        st.error("GEMINI_API_KEY가 설정되지 않았습니다. Streamlit Secrets 설정을 확인해 주세요.")
        st.stop()

    client = genai.Client(api_key=api_key)
    uploaded_gemini_files = []
    temp_paths = []

    try:
        with st.spinner("1/2. 첨부 문서 및 Claroty 자산 인벤토리 전처리 중..."):
            asset_summary_text = summarize_claroty_assets(asset_excel)

            for uploaded in [rfp_file, asis_file]:
                if uploaded is not None:
                    suffix = os.path.splitext(uploaded.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(uploaded.getvalue())
                        temp_paths.append(tmp.name)
                    g_file = client.files.upload(file=temp_paths[-1])
                    uploaded_gemini_files.append(g_file)

        with st.spinner("2/2. 자산 세그멘테이션 및 편집 가능한 PowerPoint 슬라이드 생성 중 (약 25~45초 소요)..."):
            prompt = f"""
            제공된 자료와 아래 선택 조건을 바탕으로 IT/OT 보안 To-Be 아키텍처를 설계해 주세요.
            - 적용 규제 및 표준: {', '.join(regulations) if regulations else 'IEC 62443 및 ISO 27001 기본 기준'}
            - 현재 보유 중인 솔루션 (As-Is): {', '.join(all_current) if all_current else '정보 없음'}
            - 신규 도입 예정 솔루션 (To-Be): {', '.join(all_planned) if all_planned else '규제 충족에 필요한 표준 솔루션 자동 추천'}
            - Claroty 식별 자산 분석 결과 (IP 대역 및 용도별):
{asset_summary_text if asset_summary_text else '  (첨부된 자산 인벤토리 없음)'}
            - 추가 현장 특이사항: {extra_notes if extra_notes else '없음'}
            """
            contents = uploaded_gemini_files + [prompt]

            response = client.models.generate_content(
                model="gemini-2.5-pro",
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0.2,
                    response_mime_type="application/json",
                    response_schema=ArchitectureDesignOutput,
                ),
            )
            result: ArchitectureDesignOutput = response.parsed
            pptx_bytes = build_editable_pptx(result)

        st.success("To-Be 아키텍처 세그멘테이션 및 PowerPoint 생성이 완료되었습니다!")

        st.download_button(
            label="📊 편집 가능한 PowerPoint(.pptx) 다운로드 (요약 + 도형 구성도 + 세그멘테이션/매트릭스 3슬라이드)",
            data=pptx_bytes,
            file_name="IT_OT_ToBe_Security_Architecture.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 웹 미리보기 (Mermaid)",
            "📋 규제·세그멘테이션 매트릭스",
            "🗂️ Claroty 자산 요약 결과",
            "📝 설계 요약 및 코드"
        ])

        with tab1:
            st.subheader("자동 생성된 To-Be 보안 아키텍처 (웹 미리보기)")
            render_mermaid(result.mermaid_code)

        with tab2:
            st.subheader("규제 충족 및 Zone 세그멘테이션 정책 매트릭스")
            df = pd.DataFrame([item.model_dump() for item in result.compliance_matrix])
            df.columns = ["조항/세그먼트 구간", "보안 영역", "As-Is 취약점/혼재", "To-Be 통제/허용정책", "충족 근거"]
            st.dataframe(df, use_container_width=True)

            csv_data = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 매트릭스 엑셀(CSV) 다운로드", data=csv_data, file_name="Security_Segmentation_Matrix.csv", mime="text/csv")

        with tab3:
            st.subheader("업로드된 Claroty 자산 엑셀 IP/용도별 그룹핑 결과")
            if asset_summary_text:
                st.code(asset_summary_text, language="text")
            else:
                st.info("Claroty 자산 엑셀/CSV 파일이 업로드되지 않았습니다. 좌측 사이드바에서 파일을 올리면 IP 대역과 장비 유형별 요약 결과가 여기에 표시됩니다.")

        with tab4:
            st.subheader("핵심 설계 요약")
            st.write(result.executive_summary)
            st.subheader("분리된 보안 구역 (Zones / 등급)")
            for z in result.identified_zones:
                st.markdown(f"- {z}")
            st.subheader("Mermaid 원본 코드")
            st.code(result.mermaid_code, language="mermaid")

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")

    finally:
        for gf in uploaded_gemini_files:
            try:
                client.files.delete(name=gf.name)
            except Exception:
                pass
        for tp in temp_paths:
            if os.path.exists(tp):
                os.remove(tp)
