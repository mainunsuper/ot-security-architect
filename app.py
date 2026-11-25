import os
import tempfile
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from typing import List
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# 페이지 기본 설정
st.set_page_config(page_title="OT 보안 To-Be 아키텍트 AI", layout="wide")

# ==========================================
# 1. 팀원 전용 비밀번호 인증 기능
# ==========================================
def check_password() -> bool:
    """Streamlit Secrets의 APP_PASSWORD와 비교하여 접근을 통제합니다."""
    if st.session_state.get("authenticated", False):
        return True

    st.title("🔒 OT 보안 아키텍트 AI - 접근 제한")
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
st.title("🛡️ OT 보안 To-Be 아키텍처 자동 설계 에이전트")
st.markdown("RFP 문서, 기존 구성도, 보유/도입 솔루션을 바탕으로 **IEC 62443 / NIS2 / IAEA** 기반 To-Be 구성도와 매트릭스를 자동 생성합니다.")

class ComplianceItem(BaseModel):
    requirement_id: str = Field(description="RFP 요구사항 번호 또는 표준 조항 (예: IEC 62443-3-3 SR 5.1)")
    category: str = Field(description="보안 영역 (예: 네트워크 세분화, 자산 가시성, 원격접속 통제)")
    as_is_gap: str = Field(description="기존 As-Is 구성의 취약점 또는 미비점")
    to_be_control: str = Field(description="To-Be 구성도에 반영된 보안 통제 방안 및 솔루션")
    justification: str = Field(description="해당 설계가 요건을 충족하는 기술적 근거")

class ArchitectureDesignOutput(BaseModel):
    executive_summary: str = Field(description="고객사 환경 분석 및 To-Be 설계 핵심 요약")
    identified_zones: List[str] = Field(description="IEC 62443 기준으로 분리된 Zone 목록")
    mermaid_code: str = Field(description="렌더링 가능한 순수 Mermaid.js 코드 (```mermaid 태그 제외)")
    compliance_matrix: List[ComplianceItem] = Field(description="규제 및 RFP 요건 충족 매트릭스")

SYSTEM_PROMPT = """
당신은 15년 이상 경력의 OT/ICS 보안 수석 아키텍트입니다.
사용자가 제공하는 고객사 RFP(문서), 기존 As-Is 네트워크 구성도, 적용 규제 기준, 그리고 [기존 구성 솔루션] 및 [도입 예정 솔루션] 정보를 종합하여 최적의 To-Be 보안 아키텍처를 설계합니다.

[설계 원칙]
1. Purdue Model(Level 0~4) 계층 구조를 명확히 나누고, Level 3(OT)와 Level 4(IT) 사이에 반드시 Level 3.5(IDMZ)를 배치합니다.
2. IEC 62443 기준에 따라 기능/중요도별 Zone을 분리하고, Zone 간 통신은 산업용 방화벽 또는 일방향 망연계(Conduit)로 통제합니다.
3. 사용자가 선택한 [기존 구성 솔루션]은 As-Is 영역에서 최대한 유지/연동하되 취약 구간을 보완하고, [도입 예정 솔루션]은 각 솔루션의 표준 아키텍처 위치에 정확히 배치합니다.
4. [세부 솔루션별 표준 배치 규칙]
   - Claroty CTD / xDome / Medigate: 코어 및 주요 L2/L3 백본 스위치에서 SPAN/TAP 미러링(점선 `-.->`)으로 트래픽을 수동 수집하도록 배치합니다.
   - Claroty Edge: SPAN 포트 구성이 어려운 원격지 변전소, 소규모 PLC/RTU 구간, 또는 독립 의료기기 네트워크에 배치하여 능동(Active) 수집 후 중앙 서버(CTD/xDome)로 전송하게 그립니다.
   - Claroty SRA / xSA: 반드시 Level 3.5(IDMZ)에 배치하며, 외부 벤더나 유지보수 인력이 SRA를 거쳐 승인된 자산(PLC, HMI, 의료장비)에만 접근하도록 구성합니다.
   - 일방향 망연계(Data Diode): Level 3(OT 제어망)에서 Level 4(IT 업무망) 또는 외부 관제센터로 데이터를 보낼 때만 단방향 화살표(`-->`)로 배치합니다.
   - Trellix / 화이트리스트 백신 및 매체제어: Level 2~3의 HMI, EWS(엔지니어링 워크스테이션), SCADA 서버에 적용된 것으로 표기합니다.
5. 컴플라이언스 매트릭스에는 도입 예정 솔루션이 구체적으로 어떤 규제 조항을 해결하는지 명시합니다.

[Mermaid 코드 작성 규칙]
- 반드시 `flowchart TB` 방향을 사용하며, 각 계층과 Zone을 `subgraph`로 묶습니다.
- 노드 ID는 영문/숫자만 사용하고, 한글 라벨은 반드시 따옴표(`["한글 라벨"]`)로 감쌉니다.
"""

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
    components.html(html_code, height=650, scrolling=True)

# ==========================================
# 3. 좌측 사이드바: 자료, 규제, 솔루션 선택 및 직접 입력 UI
# ==========================================
OT_SOLUTIONS_LIST = [
    # [공통 - 자산 가시성 및 위협 탐지]
    "없음 (신규 구축)",
    "Claroty CTD (OT 수동형 자산 가시성/위협 탐지)",
    "Claroty xDome (클라우드 기반 OT/IoT 보안 플랫폼)",
    "Claroty Edge (SPAN 불가 구간 능동형 자산 수집)",
    
    # [공통 - 네트워크 경계 및 원격접속]
    "Claroty SRA / xSA (OT 전용 안전한 원격접속)",
    "산업용 방화벽 (FortiGate Rugged / Hirschmann / Stratix)",
    "일반 IT/경계 방화벽 (Palo Alto / Fortinet 등)",
    "일방향 망연계 (Data Diode - 제어망→업무망 단방향 전송)",
    "양방향 망연계 (IDMZ 중계 및 패치/파일 전송)",
    "산업용 L2/L3 보안 스위치 (Cisco IE / Hirschmann / Rockwell Stratix)",
    
    # [발전·에너지 / 기반시설 특화]
    "발전 제어망 전용 백신/화이트리스트 (Trellix / AhnLab EPS 등)",
    "매체제어 및 휴대용 저장매체(USB) 반입통제 솔루션",
    "발전소 엔지니어링 워크스테이션(EWS) 접근통제/PAM",
    
    # [병원·의료(IoMT) 특화]
    "Claroty Medigate / xDome Healthcare (의료기기 IoMT 자산·위협 관리)",
    "의료기기 전용 마이크로 세그멘테이션 (NAC/방화벽 정책 연동)",
    "PACS / EMR 연동 구간 전용 보안 게이트웨이",
    
    # [공통 - 운영 및 관제]
    "NAC (네트워크 접근제어)",
    "통합로그 / OT SIEM 연동 (Splunk / QRadar 등)"
]

with st.sidebar:
    st.header("1. 분석 자료 업로드")
    rfp_file = st.file_uploader("고객사 RFP 또는 요구사항 문서 (PDF)", type=["pdf"])
    asis_file = st.file_uploader("기존 As-Is 구성도 (PNG, JPG, PDF)", type=["png", "jpg", "jpeg", "pdf"])
    
    st.header("2. 보안 솔루션 구성 선택")
    current_solutions = st.multiselect(
        "현재 구성되어 있는 보안 솔루션 (As-Is)",
        OT_SOLUTIONS_LIST,
        default=["일반 IT/경계 방화벽 (Palo Alto / Fortinet 등)"],
        help="고객사 현장에 이미 도입되어 운영 중인 솔루션을 선택하세요."
    )
    custom_current = st.text_input(
        "기타 기존 보유 솔루션 직접 입력 (선택)", 
        placeholder="예: 시스코 ISE, 노후화 방화벽 등 (쉼표로 구분)"
    )

    st.markdown("---")

    planned_solutions = st.multiselect(
        "도입 예정인 보안 솔루션 (To-Be)",
        [s for s in OT_SOLUTIONS_LIST if s != "없음 (신규 구축)"],
        default=[
            "산업용 방화벽 (FortiGate Rugged / Hirschmann / Stratix)",
            "Claroty CTD (OT 수동형 자산 가시성/위협 탐지)",
            "Claroty SRA / xSA (OT 전용 안전한 원격접속)"
        ],
        help="이번 사업을 통해 신규 도입하거나 제안할 솔루션을 선택하세요."
    )
    custom_planned = st.text_input(
        "기타 도입 예정 솔루션 직접 입력 (선택)", 
        placeholder="예: 자사 특화 OT 관제, 신규 IDS 등 (쉼표로 구분)"
    )

    st.header("3. 규제 및 추가 요건")
    regulations = st.multiselect(
        "적용할 보안 표준 및 규제 선택",
        ["IEC 62443-3-3 (SL-2)", "IEC 62443-3-3 (SL-3)", "NIS2 Directive", "IAEA NSS 17-T (원자력/발전)", "국내 주요정보통신기반시설 보호지침"],
        default=["IEC 62443-3-3 (SL-2)", "국내 주요정보통신기반시설 보호지침"]
    )
    extra_notes = st.text_area("추가 반영 사항 (선택)", placeholder="예: 1호기~4호기 PLC 구간은 SPAN 미러링 불가로 Claroty Edge 적용 필요 등")
    run_btn = st.button("🚀 To-Be 아키텍처 생성하기", type="primary", use_container_width=True)

# ==========================================
# 4. 메인 화면 실행 로직
# ==========================================
if run_btn:
    # 드롭다운 선택 항목과 직접 입력 항목 병합
    all_current = list(current_solutions)
    if custom_current.strip():
        all_current.extend([item.strip() for item in custom_current.split(",") if item.strip()])

    all_planned = list(planned_solutions)
    if custom_planned.strip():
        all_planned.extend([item.strip() for item in custom_planned.split(",") if item.strip()])

    if not rfp_file and not asis_file and not extra_notes and not all_planned:
        st.warning("분석할 파일이나 도입 예정 솔루션, 추가 요구사항 중 최소 1가지를 입력해 주세요.")
        st.stop()

    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        st.error("GEMINI_API_KEY가 설정되지 않았습니다. Streamlit Secrets 설정을 확인해 주세요.")
        st.stop()

    client = genai.Client(api_key=api_key)
    uploaded_gemini_files = []
    temp_paths = []

    try:
        with st.spinner("1/2. 첨부 문서를 제미나이 엔진으로 전송 중..."):
            for uploaded in [rfp_file, asis_file]:
                if uploaded is not None:
                    suffix = os.path.splitext(uploaded.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(uploaded.getvalue())
                        temp_paths.append(tmp.name)
                    g_file = client.files.upload(file=temp_paths[-1])
                    uploaded_gemini_files.append(g_file)

        with st.spinner("2/2. 선택된 솔루션과 규제를 반영해 To-Be 구성도를 설계 중입니다 (약 20~40초 소요)..."):
            prompt = f"""
            제공된 자료와 아래 선택 조건을 바탕으로 OT 보안 To-Be 아키텍처를 설계해 주세요.
            - 적용 규제 및 표준: {', '.join(regulations) if regulations else 'IEC 62443 기본 기준'}
            - 현재 보유 중인 솔루션 (As-Is): {', '.join(all_current) if all_current else '정보 없음'}
            - 신규 도입 예정 솔루션 (To-Be): {', '.join(all_planned) if all_planned else '규제 충족에 필요한 표준 솔루션 자동 추천'}
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

        st.success("To-Be 아키텍처 설계가 완료되었습니다!")

        tab1, tab2, tab3 = st.tabs(["📊 To-Be 구성도", "📋 컴플라이언스 매트릭스", "📝 설계 요약 및 코드"])

        with tab1:
            st.subheader("자동 생성된 To-Be 보안 아키텍처")
            render_mermaid(result.mermaid_code)
            st.info("💡 구성도를 수정하고 싶다면 '설계 요약 및 코드' 탭에서 Mermaid 코드를 복사해 mermaid.live 사이트에서 편집할 수 있습니다.")

        with tab2:
            st.subheader("규제 및 RFP 요건 충족 매트릭스")
            df = pd.DataFrame([item.model_dump() for item in result.compliance_matrix])
            df.columns = ["조항/요구사항", "보안 영역", "As-Is 취약점", "To-Be 통제 방안", "충족 근거"]
            st.dataframe(df, use_container_width=True)

            csv_data = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 매트릭스 엑셀(CSV) 다운로드", data=csv_data, file_name="OT_Compliance_Matrix.csv", mime="text/csv")

        with tab3:
            st.subheader("핵심 설계 요약")
            st.write(result.executive_summary)
            st.subheader("분리된 보안 구역 (Zones)")
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
