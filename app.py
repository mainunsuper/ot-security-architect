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
st.title("🛡️ OT 보안 To-Be 아키텍처 자동 설계 에이전트")
st.markdown("RFP 문서와 기존 As-Is 구성도를 업로드하면 **IEC 62443 / NIS2 / IAEA** 기반의 To-Be 구성도와 컴플라이언스 매트릭스를 자동 생성합니다.")

# 출력 데이터 규격 정의
class ComplianceItem(BaseModel):
    requirement_id: str = Field(description="RFP 요구사항 번호 또는 표준 조항 (예: IEC 62443-3-3 SR 5.1)")
    category: str = Field(description="보안 영역 (예: 네트워크 세분화, 자산 가시성, 원격접속 통제)")
    as_is_gap: str = Field(description="기존 As-Is 구성의 취약점 또는 미비점")
    to_be_control: str = Field(description="To-Be 구성도에 반영된 보안 통제 방안")
    justification: str = Field(description="해당 설계가 요건을 충족하는 기술적 근거")

class ArchitectureDesignOutput(BaseModel):
    executive_summary: str = Field(description="고객사 환경 분석 및 To-Be 설계 핵심 요약")
    identified_zones: List[str] = Field(description="IEC 62443 기준으로 분리된 Zone 목록")
    mermaid_code: str = Field(description="렌더링 가능한 순수 Mermaid.js 코드 (```mermaid 태그 제외)")
    compliance_matrix: List[ComplianceItem] = Field(description="규제 및 RFP 요건 충족 매트릭스")

SYSTEM_PROMPT = """
당신은 15년 이상 경력의 OT/ICS 보안 수석 아키텍트입니다.
사용자가 제공하는 고객사 RFP(문서), 기존 As-Is 네트워크 구성도(이미지/PDF), 및 적용 규제 기준을 분석하여 최적의 To-Be 보안 아키텍처를 설계합니다.

[설계 원칙]
1. Purdue Model(Level 0~4) 계층 구조를 명확히 나누고, Level 3(OT)와 Level 4(IT) 사이에 반드시 Level 3.5(IDMZ)를 배치합니다.
2. IEC 62443 기준에 따라 기능/중요도별 Zone을 분리하고, Zone 간 통신은 산업용 방화벽 또는 일방향 망연계(Conduit)로 통제합니다.
3. 주요 백본 스위치의 SPAN/TAP 포트를 통해 OT 보안 모니터링 센서가 수동적(Passive)으로 자산과 트래픽을 수집하도록 구성합니다.
4. 외부 벤더 원격 접속은 IDMZ 내 점프 호스트 및 OT 전용 원격접속(SRA) 솔루션을 경유하도록 설계합니다.

[Mermaid 코드 작성 규칙]
- 반드시 `flowchart TB` 방향을 사용하며, 각 계층과 Zone을 `subgraph`로 묶습니다.
- 노드 ID는 영문/숫자만 사용하고, 한글 라벨은 반드시 따옴표(`["한글 라벨"]`)로 감쌉니다.
"""

# Mermaid 다이어그램을 웹에서 바로 그려주는 함수
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

# 좌측 사이드바: 입력 영역
with st.sidebar:
    st.header("1. 분석 자료 업로드")
    rfp_file = st.file_uploader("고객사 RFP 또는 요구사항 문서 (PDF)", type=["pdf"])
    asis_file = st.file_uploader("기존 As-Is 구성도 (PNG, JPG, PDF)", type=["png", "jpg", "jpeg", "pdf"])
    
    st.header("2. 규제 및 산업 요건")
    regulations = st.multiselect(
        "적용할 보안 표준 및 규제 선택",
        ["IEC 62443-3-3 (SL-2)", "IEC 62443-3-3 (SL-3)", "NIS2 Directive", "IAEA NSS 17-T (원자력/발전)", "국내 주요정보통신기반시설 보호지침"],
        default=["IEC 62443-3-3 (SL-2)", "국내 주요정보통신기반시설 보호지침"]
    )
    extra_notes = st.text_area("추가 반영 사항 (선택)", placeholder="예: Claroty CTD 센서 도입 반영, 망연계는 일방향 솔루션 우선 적용 등")
    run_btn = st.button("🚀 To-Be 아키텍처 생성하기", type="primary", use_container_width=True)

# 메인 화면 실행 로직
if run_btn:
    if not rfp_file and not asis_file and not extra_notes:
        st.warning("최소 1개 이상의 파일(RFP 또는 구성도)이나 추가 요구사항을 입력해 주세요.")
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

        with st.spinner("2/2. 제미나이가 OT 인프라와 규제를 분석해 To-Be 구성도를 설계 중입니다 (약 20~40초 소요)..."):
            prompt = f"""
            제공된 자료를 바탕으로 다음 규제 기준을 충족하는 OT 보안 To-Be 아키텍처를 설계해 주세요.
            - 적용 규제: {', '.join(regulations)}
            - 추가 요청사항: {extra_notes if extra_notes else '없음'}
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

        # 결과 출력 탭 구성
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

            # 마크다운/CSV 다운로드 버튼
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
        # 임시 파일 정리
        for gf in uploaded_gemini_files:
            try:
                client.files.delete(name=gf.name)
            except Exception:
                pass
        for tp in temp_paths:
            if os.path.exists(tp):
                os.remove(tp)
