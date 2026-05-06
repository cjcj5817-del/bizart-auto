# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ── 설정 ──────────────────────────────────────────────────────────
MIN_VM      = 500_000
SUB_MONTHS  = 23
MAIL_TO1    = "artmanse21@naver.com"
MAIL_TO2    = "seouldm@seouldm.co.kr"
MAIL_FROM   = "coco5817@valuemark.co.kr"
SMTP_SERVER = "api.mail.bizbee.co.kr"
SMTP_PORT   = 587

# ── 유틸 ──────────────────────────────────────────────────────────
def add_months(yyyymm, m):
    y, mo = yyyymm // 100, yyyymm % 100
    t = y * 12 + (mo - 1) + m
    return (t // 12) * 100 + (t % 12 + 1)

def sv(v):
    return '' if v is None else str(v).strip()

def nv(v):
    try: return float(v) if v is not None else 0
    except: return 0

# ── 핵심 처리 ─────────────────────────────────────────────────────
def process_bizart(file_bytes, mak, filename):
    log = []
    now = datetime.now()
    log.append(f"실행시각: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    log.append(f"파일: {filename}")
    log.append(f"마감월: {mak}  →  최종발송호: {add_months(mak, SUB_MONTHS)}")

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), keep_vba=True)

    # 참조 데이터 로드
    key_map = {}
    for row in wb['키'].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            key_map[sv(row[0])] = int(row[1]) if row[1] is not None else 0

    addr_map = {}
    for row in wb['업체주소'].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            addr_map[sv(row[0])] = sv(row[3]) if len(row) > 3 else ''

    hr_map = {}
    for row in wb['인사'].iter_rows(min_row=4, values_only=True):
        if len(row) > 1 and row[1] is not None:
            hr_map[sv(row[1])] = sv(row[40]) if len(row) > 40 else ''

    log.append(f"참조 데이터 — 키: {len(key_map)}개 | 업체주소: {len(addr_map)}개 | 인사: {len(hr_map)}명")

    # 장기발송명단(신규) 처리
    ws = wb['장기발송명단(신규)']
    total = ok = ex_cont = ex_amt = no_addr = na_send = 0
    entries = []

    for ri in range(5, ws.max_row + 1):
        if ws.cell(ri, 8).value is None:
            continue
        계약상태  = sv(ws.cell(ri, 16).value)
        vm        = nv(ws.cell(ri, 30).value)
        증권      = sv(ws.cell(ri,  8).value)
        사번      = sv(ws.cell(ri,  3).value)
        모집      = sv(ws.cell(ri,  2).value)
        계약자    = sv(ws.cell(ri,  9).value)

        유지      = '대상' if key_map.get(계약상태, 0) == 1 else '제외'
        금액      = '대상' if vm >= MIN_VM else '제외'
        받는주소  = addr_map.get(증권, '')
        주소유무  = '대상' if 받는주소.strip() else '제외'
        최종      = '대상' if (유지 == '대상' and 금액 == '대상') else '제외'
        보내는주소 = hr_map.get(사번, '#N/A')
        업체명    = f"{계약자} 대표님" if 계약자 else ''

        ws.cell(ri, 76).value = 유지
        ws.cell(ri, 77).value = 금액
        ws.cell(ri, 78).value = 주소유무
        ws.cell(ri, 79).value = 최종
        ws.cell(ri, 80).value = mak
        ws.cell(ri, 81).value = add_months(mak, SUB_MONTHS)
        ws.cell(ri, 82).value = 사번
        ws.cell(ri, 83).value = 모집
        ws.cell(ri, 84).value = 보내는주소
        ws.cell(ri, 85).value = 업체명
        ws.cell(ri, 86).value = 받는주소
        ws.cell(ri, 87).value = '중기이코노미기업지원단'
        ws.cell(ri, 88).value = 1
        ws.cell(ri, 89).value = None
        ws.cell(ri, 90).value = None
        ws.cell(ri, 91).value = 증권

        total += 1
        if 최종 == '대상':
            ok += 1
            entries.append({'사번': 사번, '모집': 모집, '보내는주소': 보내는주소,
                            '업체명': 업체명, '받는주소': 받는주소, '증권': 증권,
                            '최초': mak, '최종': add_months(mak, SUB_MONTHS)})
        else:
            if 유지 == '제외': ex_cont += 1
            if 금액 == '제외': ex_amt  += 1
        if not 받는주소.strip(): no_addr += 1
        if 보내는주소 == '#N/A': na_send += 1

    log.append(f"신규 처리 — 전체: {total}건 | 최종대상: {ok}건 | 계약상태제외: {ex_cont}건 | 금액미달제외: {ex_amt}건")

    # (서울DM) 장기발송명단 업데이트
    dm   = wb['(서울DM) 장기발송명단']
    last = dm.max_row
    while last >= 6:
        if dm.cell(last, 4).value is not None: break
        last -= 1

    exist = set()
    for r in range(6, last + 1):
        v = dm.cell(r, 13).value
        if v: exist.add(sv(v))

    added = dup = no_a = 0
    for e in entries:
        bp = sv(e['보내는주소'])
        rp = sv(e['받는주소'])
        zp = sv(e['증권'])
        if not bp or bp == '#N/A': no_a += 1; continue
        if not rp:                  no_a += 1; continue
        if zp in exist:             dup  += 1; continue
        nr = last + 1 + added
        dm.cell(nr,  2).value = e['최초']
        dm.cell(nr,  3).value = e['최종']
        dm.cell(nr,  4).value = e['사번']
        dm.cell(nr,  5).value = e['모집']
        dm.cell(nr,  6).value = bp
        dm.cell(nr,  7).value = e['업체명']
        dm.cell(nr,  8).value = rp
        dm.cell(nr,  9).value = '중기이코노미기업지원단'
        dm.cell(nr, 10).value = 1
        dm.cell(nr, 11).value = None
        dm.cell(nr, 12).value = None
        dm.cell(nr, 13).value = zp
        exist.add(zp); added += 1

    log.append(f"발송명단 업데이트 — 추가: {added}건 | 중복제외: {dup}건 | 주소없어제외: {no_a}건")

    # 결과 xlsm
    buf_xlsm = io.BytesIO()
    wb.save(buf_xlsm)
    buf_xlsm.seek(0)

    # 첨부용 xlsx
    m2  = re.search(r'(\d{2})\.(\d{2})월호',   filename)
    m3  = re.search(r'(\d{2})\.(\d{2})월마감', filename)
    wyy = m2.group(1) if m2 else str(now.year)[2:]
    wmm = m2.group(2) if m2 else f"{now.month:02d}"
    myy = m3.group(1) if m3 else wyy
    mmm = m3.group(2) if m3 else f"{(now.month-2)%12+1:02d}"
    attach_name = f"{wyy}.{wmm}월호 비자트 ({myy}.{mmm}마감).xlsx"

    wb_att = openpyxl.Workbook()
    ws_att = wb_att.active
    ws_att.title = "발송명단"
    for ci, h in enumerate(['최초발송호','최종발송호','사번','보내는사람','보내는주소',
                             '업체명','받는주소','분류','중기이코노미기업지원단','증권번호'], 1):
        ws_att.cell(1, ci).value = h
    row_out = 2
    dm2 = wb['(서울DM) 장기발송명단']
    for ri2 in range(6, dm2.max_row + 1):
        if dm2.cell(ri2, 4).value is None: continue
        for ci, v in enumerate([dm2.cell(ri2, c).value for c in [2,3,4,5,6,7,8,9,10,13]], 1):
            ws_att.cell(row_out, ci).value = v
        row_out += 1
    buf_xlsx = io.BytesIO()
    wb_att.save(buf_xlsx)
    buf_xlsx.seek(0)

    log.append(f"첨부 엑셀 생성 — {attach_name} ({row_out-2}건)")

    # 메일 내용
    send_date    = now.strftime('%Y.%m.%d')
    mail_subject = f"(주)밸류마크 {wyy}.{wmm}월호 비자트 발주 내용 전달 ({send_date})"
    mail_body    = f"안녕하세요. 밸류마크 총무팀입니다.\n{wyy}년 {wmm}월호 비자트 발주 내용 전달드립니다.\n감사합니다."

    stats = dict(
        total=total, ok=ok, added=added, dup=dup, no_a=no_a,
        na_send=na_send, ex_cont=ex_cont, ex_amt=ex_amt, no_addr=no_addr,
        attach_name=attach_name, mail_subject=mail_subject,
        mail_body=mail_body, log=log,
        wyy=wyy, wmm=wmm
    )
    return buf_xlsm.read(), buf_xlsx.read(), stats

# ── SMTP 발송 ──────────────────────────────────────────────────────
def send_mail(mail_pw, subject, body, attach_bytes, attach_name):
    msg = MIMEMultipart()
    msg['From']    = MAIL_FROM
    msg['To']      = f"{MAIL_TO1}, {MAIL_TO2}"
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attach_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', attach_name))
    msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15) as server:
        server.ehlo(); server.starttls(); server.ehlo()
        server.login(MAIL_FROM, mail_pw)
        server.sendmail(MAIL_FROM, [MAIL_TO1, MAIL_TO2], msg.as_bytes())

# ══════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="비자트 자동화", page_icon="📮", layout="wide")

# 커스텀 CSS
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f8fafc; }
[data-testid="stSidebar"] { background: #1e3a5f; }
[data-testid="stSidebar"] * { color: white !important; }
.login-box {
    max-width: 400px; margin: 80px auto; padding: 40px;
    background: white; border-radius: 16px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.10);
}
.step-box {
    background: white; border-radius: 12px; padding: 20px 24px;
    margin-bottom: 16px; border-left: 4px solid #2563eb;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-card {
    background: white; border-radius: 10px; padding: 16px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}
</style>
""", unsafe_allow_html=True)

# ── 로그인 ────────────────────────────────────────────────────────
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.markdown("""
    <div class="login-box">
        <h2 style="text-align:center; color:#1e3a5f; margin-bottom:8px;">📮 비자트 자동화</h2>
        <p style="text-align:center; color:#888; margin-bottom:24px; font-size:14px;">(주)밸류마크 총무팀</p>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        pw = st.text_input("접속 비밀번호를 입력하세요", type="password", label_visibility="collapsed",
                           placeholder="비밀번호")
        if st.button("로그인", use_container_width=True, type="primary"):
            if pw == st.secrets.get("app_password", "valuemark2026"):
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

# ── 사이드바 ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📮 비자트 자동화")
    st.markdown("**(주)밸류마크 총무팀**")
    st.divider()
    st.markdown("### 📋 작업 순서")
    st.markdown("""
1. xlsm 파일 업로드
2. 마감월 확인
3. 비밀번호 입력 *(선택)*
4. **자동화 실행** 버튼 클릭
5. 파일 다운로드 & 메일 발송
    """)
    st.divider()
    st.markdown("### ✅ 자동 처리 항목")
    st.markdown("""
- 계약유지 여부 확인
- VM환산 금액 충족 확인 (50만원↑)
- 업체주소 매핑
- 인사정보 매핑
- 장기발송명단 업데이트
- 발주 첨부 엑셀 생성
- 발주 메일 초안 생성
    """)
    st.divider()
    if st.button("🔒 로그아웃"):
        st.session_state.auth = False
        st.session_state.processed = False
        st.rerun()

# ── 메인 ─────────────────────────────────────────────────────────
st.title("📮 비자트 발송 자동화")
st.caption(f"실행일: {datetime.now().strftime('%Y년 %m월 %d일')}  |  담당: 총무팀")
st.divider()

# 업로드 & 설정
col_up, col_cfg = st.columns([3, 2])

with col_up:
    uploaded = st.file_uploader(
        "📁 비자트 xlsm 파일 업로드",
        type=["xlsm"],
        help="예: 26.05월호 비자트 (26.03월마감).xlsm"
    )

with col_cfg:
    now = datetime.now()
    auto_mak = None
    if uploaded:
        m = re.search(r'(\d{2})\.(\d{2})월마감', uploaded.name)
        if m:
            auto_mak = (2000 + int(m.group(1))) * 100 + int(m.group(2))
    default_mak = str(auto_mak) if auto_mak else f"{now.year}{now.month:02d}"

    mak_input = st.text_input(
        "📅 마감월",
        value=default_mak,
        help="파일명에서 자동 감지됩니다. 형식: YYYYMM (예: 202603)"
    )
    mail_pw_input = st.text_input(
        "🔑 비즈비웍스 비밀번호 (메일 자동발송용)",
        type="password",
        help="입력 시 처리 완료 후 예술만세·서울DM으로 자동 발송됩니다. 비워두면 메일 초안만 표시됩니다."
    )

# 실행 버튼
if not uploaded:
    st.info("위에서 xlsm 파일을 업로드하면 자동화 실행 버튼이 활성화됩니다.")
else:
    if st.button("🚀 자동화 실행", type="primary", use_container_width=True):
        try:
            mak = int(mak_input)
        except:
            st.error("마감월 형식이 올바르지 않습니다. 예시: 202603")
            st.stop()

        with st.spinner("⏳ 처리 중입니다. 잠시만 기다려주세요..."):
            file_bytes = uploaded.read()
            xlsm_out, xlsx_out, stats = process_bizart(file_bytes, mak, uploaded.name)

        st.session_state.update({
            "xlsm_out":  xlsm_out,
            "xlsx_out":  xlsx_out,
            "stats":     stats,
            "mail_pw":   mail_pw_input,
            "processed": True,
            "filename":  uploaded.name,
            "mail_sent": False,
        })

        # 비밀번호 있으면 즉시 자동 발송
        if mail_pw_input:
            with st.spinner("📨 메일 발송 중..."):
                try:
                    send_mail(mail_pw_input, stats["mail_subject"], stats["mail_body"],
                              xlsx_out, stats["attach_name"])
                    st.session_state["mail_sent"] = True
                except smtplib.SMTPAuthenticationError:
                    st.session_state["mail_error"] = "로그인 실패 — 비밀번호를 확인해주세요."
                except Exception as e:
                    st.session_state["mail_error"] = str(e)

# ── 결과 화면 ─────────────────────────────────────────────────────
if st.session_state.get("processed"):
    stats       = st.session_state["stats"]
    xlsm_out    = st.session_state["xlsm_out"]
    xlsx_out    = st.session_state["xlsx_out"]
    attach_name = stats["attach_name"]
    filename    = st.session_state["filename"]
    mail_sent   = st.session_state.get("mail_sent", False)
    mail_error  = st.session_state.get("mail_error", "")

    st.divider()

    # 완료 메시지
    if mail_sent:
        st.success(f"✅ 자동화 완료!  |  📨 메일 발송 완료 ({MAIL_TO1}, {MAIL_TO2})")
    else:
        st.success("✅ 자동화 완료!")

    # 오류 메시지
    if mail_error:
        st.error(f"❌ 메일 발송 실패: {mail_error}")

    # 핵심 지표
    st.markdown("### 📊 처리 결과")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("신규 계약", f"{stats['total']}건")
    c2.metric("최종 발송 대상", f"{stats['ok']}건")
    c3.metric("발송명단 신규 추가", f"{stats['added']}건")
    c4.metric("중복으로 제외", f"{stats['dup']}건")
    c5.metric("주소 없어 제외", f"{stats['no_a']}건")

    # 확인 필요 알림
    if stats['na_send'] > 0:
        st.warning(f"⚠️ 보내는 주소 없음 {stats['na_send']}건 — 인사 시트를 최신 데이터로 업데이트하거나 직접 입력이 필요합니다.")
    if stats['no_addr'] > 0:
        st.info(f"ℹ️ 받는 주소 없음 {stats['no_addr']}건 — 업체주소 시트에 추가하면 다음 달 자동으로 포함됩니다.")

    st.divider()

    # 파일 다운로드
    st.markdown("### 📥 파일 다운로드")
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="📊 결과 xlsm 다운로드 (업데이트된 비자트 파일)",
            data=xlsm_out,
            file_name=filename,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True
        )
    with dl2:
        st.download_button(
            label="📎 발주 첨부용 xlsx 다운로드",
            data=xlsx_out,
            file_name=attach_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.divider()

    # 발주 메일
    st.markdown("### 📧 발주 메일")

    if mail_sent:
        st.success(f"메일이 자동 발송되었습니다. 수신자가 받지 못한 경우 아래 내용으로 직접 발송하세요.")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown("**수신인**")
        st.markdown(f"- 예술만세: `{MAIL_TO1}`")
        st.markdown(f"- 서울DM: `{MAIL_TO2}`")
    with col_m2:
        st.markdown("**첨부파일**")
        st.markdown(f"📎 `{attach_name}`")

    st.markdown("**제목** (우측 아이콘으로 복사)")
    st.code(stats["mail_subject"], language=None)

    st.markdown("**본문** (우측 아이콘으로 복사)")
    st.code(stats["mail_body"], language=None)

    # 비밀번호 미입력 시 수동 발송 버튼 제공
    if not mail_sent and not st.session_state.get("mail_pw"):
        st.markdown("---")
        st.markdown("**메일 자동 발송** — 비밀번호를 입력하고 발송하세요.")
        pw2 = st.text_input("비즈비웍스 비밀번호", type="password", key="pw2")
        if st.button("📨 지금 메일 발송", type="primary"):
            if pw2:
                with st.spinner("발송 중..."):
                    try:
                        send_mail(pw2, stats["mail_subject"], stats["mail_body"],
                                  xlsx_out, attach_name)
                        st.success(f"✅ 메일 발송 완료! 수신: {MAIL_TO1}, {MAIL_TO2}")
                    except smtplib.SMTPAuthenticationError:
                        st.error("❌ 로그인 실패 — 비밀번호를 확인해주세요.")
                    except Exception as e:
                        st.error(f"❌ 발송 실패: {e}")
            else:
                st.warning("비밀번호를 입력해주세요.")

    # 처리 로그
    with st.expander("📋 상세 처리 로그 보기"):
        for line in stats["log"]:
            st.text(line)
