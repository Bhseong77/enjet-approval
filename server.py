# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ENJET 메일 서버 + 결재 시스템 (v2.0)
# Last update: 2026-05-22 - Approval routes activated
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

from flask import Flask, jsonify, send_file, request, Response
import urllib.request
import urllib.parse
import urllib.error
import json
import imaplib
import email
from email.header import decode_header
from email import policy
import os
import io
import re
from datetime import datetime, timedelta

app = Flask(__name__)

# CORS - 모든 요청에 헤더 추가
@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, Range, Prefer"
    response.headers["Access-Control-Max-Age"] = "3600"
    return response

@app.errorhandler(Exception)
def handle_exception(e):
    resp = jsonify({"error": str(e)})
    resp.status_code = 500
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, Range, Prefer"
    return resp

@app.route("/api/<path:path>", methods=["OPTIONS"])
def cors_preflight(path):
    resp = Response("", 204)
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, Range, Prefer"
    resp.headers["Access-Control-Max-Age"] = "3600"
    return resp

IMAP_SERVER = os.environ.get("IMAP_SERVER", "gw.enjet.co.kr")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
EMAIL_USER = os.environ.get("EMAIL_USER", "")
EMAIL_PASS = os.environ.get("EMAIL_PASS", "")
OPENAI_KEY = os.environ.get("OPENAI_API_KEY", "")

def decode_str(s):
    if s is None: return ""
    decoded = decode_header(s)
    result = ""
    for part, enc in decoded:
        if isinstance(part, bytes):
            result += part.decode(enc or "utf-8", errors="ignore")
        else:
            result += str(part)
    return result

def parse_addresses(addr_str):
    if not addr_str: return []
    results = []
    for name, addr in email.utils.getaddresses([addr_str]):
        try: name = decode_str(name)
        except: pass
        results.append({"name": name.strip(), "email": addr.strip()})
    return results

def html_to_text(html):
    """HTML을 일반 텍스트로 변환"""
    text = re.sub(r'<style[^>]*>[\s\S]*?</style>', '', html, flags=re.I)
    text = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', text, flags=re.I)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.I)
    text = re.sub(r'</p>', '\n', text, flags=re.I)
    text = re.sub(r'</div>', '\n', text, flags=re.I)
    text = re.sub(r'<[^>]+>', '', text)
    text = text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"')
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def extract_body(msg):
    """이메일에서 본문 추출 (모든 케이스 처리)"""
    plain = ""
    html = ""

    for part in msg.walk():
        ct = part.get_content_type()
        cd = str(part.get("Content-Disposition", ""))
        if "attachment" in cd.lower():
            continue
        
        charset = part.get_content_charset() or "utf-8"
        
        if ct == "text/plain" and not plain:
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    plain = payload.decode(charset, errors="ignore")
            except: pass
        
        elif ct == "text/html" and not html:
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    html = payload.decode(charset, errors="ignore")
            except: pass
        
        elif ct == "message/rfc822":
            # 전달된 메일 처리
            try:
                fwd_msg = part.get_payload(decode=False)
                if isinstance(fwd_msg, list) and len(fwd_msg) > 0:
                    fwd_body = extract_body(fwd_msg[0])
                    if fwd_body and not plain:
                        plain = "--- 전달된 메일 ---\n" + fwd_body
            except: pass

    if plain:
        return plain.strip()
    elif html:
        return html_to_text(html)
    return ""

def extract_attachments(msg):
    """첨부파일 목록 추출 (인덱스 포함 - 다운로드 시 식별용)"""
    attachments = []
    idx = 0
    for part in msg.walk():
        cd = str(part.get("Content-Disposition", ""))
        ct = part.get_content_type()
        if "attachment" in cd.lower():
            fname = decode_str(part.get_filename() or f"attachment_{idx}")
            attachments.append({"name": fname, "content_type": ct, "idx": idx})
            idx += 1
        elif ct.startswith("image/") and "inline" in cd.lower():
            fname = decode_str(part.get_filename() or "")
            if fname:
                attachments.append({"name": fname, "content_type": ct, "inline": True, "idx": idx})
                idx += 1
    return attachments


def get_attachment_bytes(msg, target_idx):
    """msg에서 target_idx 번째 첨부의 (filename, content_type, bytes) 반환"""
    idx = 0
    for part in msg.walk():
        cd = str(part.get("Content-Disposition", ""))
        ct = part.get_content_type()
        is_attach = "attachment" in cd.lower()
        is_inline_img = ct.startswith("image/") and "inline" in cd.lower()
        if is_attach or is_inline_img:
            if idx == target_idx:
                fname = decode_str(part.get_filename() or f"attachment_{idx}")
                payload = part.get_payload(decode=True) or b""
                return fname, ct, payload
            idx += 1
    return None, None, None

# ── 기존 IMAP 엔드포인트 ────────────────────────────
def get_emails(limit=50):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("INBOX")
        _, data = mail.search(None, "ALL")
        ids = data[0].split()
        ids = ids[-limit:][::-1]
        emails = []
        for uid in ids:
            _, msg_data = mail.fetch(uid, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject = decode_str(msg.get("Subject", ""))
            sender = decode_str(msg.get("From", ""))
            cc_raw = msg.get("Cc", "") or ""
            to_raw = msg.get("To", "") or ""
            date_str = msg.get("Date", "")
            try:
                date = email.utils.parsedate_to_datetime(date_str)
                date_formatted = date.strftime("%Y-%m-%d %H:%M")
            except:
                date_formatted = date_str
            body = extract_body(msg)
            attachments = extract_attachments(msg)
            emails.append({
                "id": uid.decode(),
                "subject": subject,
                "sender": sender,
                "to": parse_addresses(to_raw),
                "cc": parse_addresses(cc_raw),
                "date": date_formatted,
                "body": body[:1000],
                "has_attachment": len(attachments) > 0,
                "attachments": attachments
            })
        mail.logout()
        return emails
    except Exception as e:
        return {"error": str(e)}

@app.route("/api/emails")
def api_emails():
    return jsonify(get_emails(50))

@app.route("/api/email/<mail_id>/attachment/<int:att_index>")
def download_attachment(mail_id, att_index):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("INBOX")
        _, msg_data = mail.fetch(mail_id.encode(), "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        mail.logout()
        att_count = 0
        for part in msg.walk():
            cd = str(part.get("Content-Disposition", ""))
            if "attachment" in cd.lower():
                if att_count == att_index:
                    fname = decode_str(part.get_filename() or f"attachment_{att_index}")
                    payload = part.get_payload(decode=True)
                    ct = part.get_content_type() or "application/octet-stream"
                    return send_file(io.BytesIO(payload), mimetype=ct, as_attachment=True, download_name=fname)
                att_count += 1
        return jsonify({"error": "첨부파일을 찾을 수 없습니다"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

from flask import send_file
import io

# ── 첨부파일 다운로드 ───────────────────────────────
@app.route("/api/eml/attachment", methods=["POST"])
def eml_attachment_endpoint():
    """
    SharePoint eml의 특정 첨부파일 다운로드
    Body: eml 바이트 (raw)
    Query: ?idx=0 (첨부 인덱스)
    """
    try:
        eml_bytes = request.get_data()
        if not eml_bytes:
            return jsonify({"error": "eml 데이터가 없습니다"}), 400
        idx = int(request.args.get("idx", "0"))

        msg = email.message_from_bytes(eml_bytes)
        fname, ct, data = get_attachment_bytes(msg, idx)
        if data is None:
            return jsonify({"error": f"첨부 #{idx} 없음"}), 404

        # 다운로드용 응답
        return send_file(
            io.BytesIO(data),
            mimetype=ct or "application/octet-stream",
            as_attachment=True,
            download_name=fname or f"attachment_{idx}"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/imap/attachment", methods=["POST"])
def imap_attachment_endpoint():
    """
    IMAP 메일의 특정 첨부파일 다운로드
    Body JSON: { user, pass, uid, folder, idx }
    """
    try:
        data = request.get_json()
        user = data.get("user", "")
        password = data.get("pass", "")
        uid = data.get("uid", "")
        folder = data.get("folder", "inbox")
        idx = int(data.get("idx", 0))
        server = data.get("server", None)

        if not all([user, password, uid]):
            return jsonify({"error": "필수 파라미터 없음"}), 400

        # IMAP에서 본문 가져오기 (안전한 5단계 fallback 사용)
        try:
            raw_bytes = _imap_fetch_raw_body_safe(user, password, server, folder, uid)
        except Exception as fetch_err:
            return jsonify({"error": f"IMAP fetch 실패: {str(fetch_err)[:200]}"}), 500

        if not raw_bytes:
            return jsonify({"error": "빈 응답"}), 500

        msg = email.message_from_bytes(raw_bytes)
        fname, ct, attach_data = get_attachment_bytes(msg, idx)
        if attach_data is None:
            return jsonify({"error": f"첨부 #{idx} 없음"}), 404

        return send_file(
            io.BytesIO(attach_data),
            mimetype=ct or "application/octet-stream",
            as_attachment=True,
            download_name=fname or f"attachment_{idx}"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── 신규: eml 파싱 엔드포인트 ───────────────────────
@app.route("/api/parse-eml", methods=["POST"])
def parse_eml_endpoint():
    """
    SharePoint에서 받아온 eml 바이트를 파싱해서 본문/첨부파일 반환
    대시보드에서 POST로 eml 파일 내용(bytes) 전송
    """
    try:
        eml_bytes = request.get_data()
        if not eml_bytes:
            return jsonify({"error": "eml 데이터가 없습니다"}), 400

        msg = email.message_from_bytes(eml_bytes)

        subject = decode_str(msg.get("Subject", ""))
        from_addr = decode_str(msg.get("From", ""))
        to_raw = msg.get("To", "") or ""
        cc_raw = msg.get("Cc", "") or ""
        date_str = msg.get("Date", "")

        try:
            date = email.utils.parsedate_to_datetime(date_str)
            date_formatted = date.strftime("%Y-%m-%d %H:%M")
        except:
            date_formatted = date_str

        body = extract_body(msg)
        attachments = extract_attachments(msg)

        return jsonify({
            "subject": subject,
            "from": from_addr,
            "to": parse_addresses(to_raw),
            "cc": parse_addresses(cc_raw),
            "date": date_formatted,
            "dateRaw": date_str,
            "body": body,
            "hasAttachment": len(attachments) > 0,
            "attachments": attachments
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})

# ── OpenAI 연락처 분석 캐시 ──────────────────────────
_ai_cache = {}

def ai_extract_contact(domain, body_text, from_addr):
    """OpenAI로 도메인+본문에서 연락처 정보 추출"""
    cache_key = domain + "|" + body_text[:100]
    if cache_key in _ai_cache:
        return _ai_cache[cache_key]

    prompt = f"""다음 이메일 본문에서 발신자의 연락처 정보를 JSON으로 추출해주세요.

이메일 도메인: {domain}
발신자 주소: {from_addr}

이메일 본문:
{body_text}

다음 형식의 JSON만 반환하세요 (설명 없이):
{{
  "name": "한글이름",
  "company": "정확한 회사명",
  "dept": "부서명 직함",
  "mobile": "휴대폰번호",
  "tel": "직통전화번호"
}}

=== 추출 규칙 ===

[이름]
- 서명의 "홍 길 동  Gildong Hong 과장" → "홍길동" (공백제거)
- "홍길동입니다", "홍길동 드림" 패턴에서 추출
- 영문 이름만 있으면 영문 그대로

[회사명] ← 가장 중요!
- 반드시 본문/서명에서 실제 회사명을 찾을 것. 도메인만으로 추측 금지.
- 삼성 계열사 구분 (samsung.com 도메인이라도 본문에서 정확히 구분):
  * 삼성전자(주) / Samsung Electronics → "삼성전자"
  * 삼성전기(주) / Samsung Electro-Mechanics / SEMCO → "삼성전기"
  * 삼성디스플레이(주) / Samsung Display / SDC → "삼성디스플레이"
  * 삼성SDI(주) / Samsung SDI → "삼성SDI"
  * 삼성물산(주) / Samsung C&T → "삼성물산"
  * 삼성SDS(주) / Samsung SDS → "삼성SDS"
  * 삼성생명(주) / Samsung Life → "삼성생명"
  * 삼성화재(주) / Samsung Fire → "삼성화재"
  * 삼성증권(주) / Samsung Securities → "삼성증권"
  * 삼성바이오로직스 / Samsung Biologics → "삼성바이오로직스"
- LG 계열사도 동일하게 본문에서 구분:
  * LG전자, LG화학, LG디스플레이, LG이노텍, LG CNS 등
- 일반 규칙: 본문에서 "XX주식회사", "XX(주)", "XX Co.,Ltd" 패턴 찾기
- 본문에 회사명이 없으면 도메인으로 추정:
  * gachon.ac.kr → 가천대학교
  * kopti.re.kr → 한국생산기술연구원
  * etri.re.kr → ETRI(한국전자통신연구원)
  * kist.re.kr → KIST(한국과학기술연구원)
  * 그 외 알 수 없으면 도메인 첫번째 파트를 회사명으로

[부서/직함]
- 서명 라인에서 추출: "설계기술팀 책임연구원", "영업1팀 과장" 등
- 직함만: 사원/주임/대리/과장/차장/부장/팀장/수석/책임/선임/이사/상무/전무/대표
- 본문 첫줄 "기구공정기술2G 서보범입니다" → dept: "기구공정기술2G"
- 부서+직함 조합: "설계팀 과장" (부서명과 직함 사이 공백 하나)
- 불필요한 것 제거: 주소, 이메일주소, 전화번호, 회사명, 영문주소

[전화번호]
- M/Mobile → mobile 필드
- T/Tel/D/Direct → tel 필드
- +82.10.xxxx.xxxx 형식도 인식
- 정보 없으면 빈 문자열 (절대 추측 금지)
"""

    try:
        payload = json.dumps({
            "model": "gpt-4o-mini",
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 300,
            "temperature": 0
        }).encode('utf-8')

        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {OPENAI_KEY}"
            }
        )
        with urllib.request.urlopen(req, timeout=15) as r:
            result = json.loads(r.read())
            text = result["choices"][0]["message"]["content"].strip()
            if "```" in text:
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            parsed = json.loads(text.strip())
            _ai_cache[cache_key] = parsed
            return parsed
    except Exception as e:
        print(f"OpenAI 오류: {e}")
        return {}


@app.route("/api/parse-eml-contact-ai", methods=["POST"])
def parse_eml_contact_ai():
    """OpenAI로 eml에서 연락처 정보 추출"""
    try:
        if not OPENAI_KEY:
            return jsonify({"error": "OPENAI_API_KEY 미설정"}), 400

        eml_bytes = request.get_data()
        if not eml_bytes:
            return jsonify({"error": "데이터 없음"}), 400

        msg = email.message_from_bytes(eml_bytes)
        from_addr = decode_str(msg.get("From", ""))
        from_email = email.utils.parseaddr(from_addr)[1]
        domain = from_email.split("@")[1] if "@" in from_email else ""
        body = extract_body(msg)

        if not body and not domain:
            return jsonify({"error": "본문 없음"}), 400

        # 본문 전체 전송 (최대 1200자) - 서명뿐 아니라 첫줄 소개도 포함
        body_text = body[:1200].strip()

        # OpenAI 분석
        result = ai_extract_contact(domain, body_text, from_addr)
        result["email"] = from_email
        result["from"] = from_addr

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── 서명에서 연락처 추출 ─────────────────────────────
TITLES = r'사원|인턴|주임|연구원|대리|과장|차장|부장|팀장|수석|책임|선임|이사|상무|전무|부사장|사장|대표이사|대표|매니저|Manager|Director|Engineer|Senior|Analyst'
DEPTS  = r'[가-힣A-Za-z]{2,15}(?:팀|부|실|센터|본부|연구소|그룹|파트|Division|Team|Dept)'

def extract_contact_from_body(body, from_addr):
    """본문 서명에서 연락처 정보 추출"""
    if not body:
        return {}

    lines = body.split('\n')
    
    # 서명 영역: 구분선 이후 또는 마지막 40줄
    sig_start = max(0, len(lines) - 40)
    for i in range(len(lines)-1, max(0, len(lines)-60), -1):
        if re.match(r'^[-_=]{2,}\s*$', lines[i].strip()) or lines[i].strip() == '--':
            sig_start = i + 1
            break
    sig_lines = lines[sig_start:]
    sig = '\n'.join(sig_lines)

    name = ""
    title = ""
    dept = ""
    mobile = ""
    tel = ""
    company = ""

    for line in sig_lines:
        line = line.strip()
        if not line:
            continue

        # ── 이름+직함+부서 라인 ──
        # "우은영  Eunyoung Woo 대리 | 기술연구소"
        # "홍길동 과장 | 영업팀"
        # "안 희 범    기업금융센터 선임매니저"
        m = re.match(
            rf'^([가-힣]{{1,2}}\s[가-힣]{{1,2}}|[가-힣]{{2,4}})\s{{1,}}'
            rf'(?:[A-Za-z\s\-]+\s+)?({TITLES})'
            rf'(?:\s*[|｜]\s*({DEPTS}))?',
            line
        )
        if m and not name:
            name = m.group(1).replace(' ', '')
            title = m.group(2)
            dept = m.group(3) or ""

        # 부서만 있는 라인 (이름 없는 경우)
        if not dept:
            m2 = re.match(rf'^({DEPTS})\s+({TITLES})', line)
            if m2:
                dept = m2.group(1)
                title = m2.group(2)

        # ── 전화번호 라인 ──
        # "M +82.10.8222.6372   T +82.70.4892.8100   F +82.70.4892.8121"
        if not mobile:
            m = re.search(r'(?:^|\|)\s*M\s+(\+?[\d.\-\s]{9,20})', line)
            if not m:
                m = re.search(r'(?:Mobile|H\.P|Cell|HP)\s*[:|]?\s*(\+?[\d.\-\s]{9,20})', line, re.I)
            if m:
                mobile = m.group(1).strip().rstrip()

        if not tel:
            m = re.search(r'(?:^|\|)\s*T\s+(\+?[\d.\-\s]{8,20})', line)
            if not m:
                m = re.search(r'(?:Tel|Phone|D)\s*[:|]?\s*(\+?[\d.\-\s]{8,20})', line, re.I)
            if m:
                tel = m.group(1).strip().rstrip()

        # 모바일/전화 둘 다 없으면 한국 번호 패턴
        if not mobile and not tel:
            phones = re.findall(
                r'(\+?82[\s.]?0?\d[\s.]\d{3,4}[\s.]\d{4}|0\d{1,2}[\s.\-]\d{3,4}[\s.\-]\d{4})',
                line
            )
            if phones:
                mobile = phones[0]
            if len(phones) > 1:
                tel = phones[1]

        # ── 회사명 ──
        if not company:
            m = re.search(
                r'([가-힣A-Za-z\s()（）]{1,20}(?:주식회사|㈜|\(주\)))'
                r'|([가-힣A-Za-z\s]{2,25}(?:은행|증권|보험|캐피탈|카드|생명|화재|투자증권|자산운용))'
                r'|([A-Za-z][A-Za-z\s&.]{2,25}(?:Co\.,?\s*Ltd\.?|Inc\.?|Corp\.))',
                line
            )
            if m:
                company = (m.group(1) or m.group(2) or m.group(3) or "").strip()

    # From 헤더에서 이름 보완
    if not name and from_addr:
        # "홍길동 <email>" 패턴
        m = re.match(r'([가-힣]{2,4})\s*<', from_addr)
        if m:
            name = m.group(1)
        else:
            # "Hong Gildong <email>" 패턴
            m = re.match(r'"?([^"<]+)"?\s*<', from_addr)
            if m:
                name = m.group(1).strip()

    dept_full = ' '.join(filter(None, [dept, title])).strip()
    
    return {
        "name": name,
        "company": company,
        "dept": dept_full,
        "mobile": mobile.strip() if mobile else "",
        "tel": tel.strip() if tel else "",
    }


@app.route("/api/parse-eml-contact", methods=["POST"])
def parse_eml_contact():
    """eml에서 연락처 정보 추출"""
    try:
        eml_bytes = request.get_data()
        if not eml_bytes:
            return jsonify({"error": "데이터 없음"}), 400

        msg = email.message_from_bytes(eml_bytes)
        from_addr = decode_str(msg.get("From", ""))
        body = extract_body(msg)
        
        contact = extract_contact_from_body(body, from_addr)
        
        # 이메일 추출
        email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', body)
        from_email = email.utils.parseaddr(from_addr)[1]
        contact["email"] = email_match.group(0).lower() if email_match else from_email
        contact["from"] = from_addr

        return jsonify(contact)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── 팀원별 IMAP 메일 조회 ────────────────────────────
def imap_connect(user, password, server=None, port=993):
    """IMAP 연결 - 한글/특수문자 비밀번호 지원"""
    import base64, socket
    srv = server or IMAP_SERVER

    # 먼저 일반 login 시도
    try:
        mail = imaplib.IMAP4_SSL(srv, port)
        mail.socket().settimeout(30)  # 30초 타임아웃
        mail.login(user, password)
        return mail
    except Exception as e1:
        err1 = str(e1)

    # AUTHENTICATE PLAIN 방식 (한글/특수문자)
    try:
        mail = imaplib.IMAP4_SSL(srv, port)
        mail.socket().settimeout(30)
        auth_str = "\x00{}\x00{}".format(user, password)
        auth_b64 = base64.b64encode(auth_str.encode("utf-8")).decode("ascii")
        typ, dat = mail._simple_command("AUTHENTICATE", "PLAIN", auth_b64)
        mail.state = "AUTH"
        if typ != "OK":
            raise imaplib.IMAP4.error("PLAIN 실패: {}".format(dat))
        return mail
    except Exception as e2:
        raise Exception("로그인 실패 (login:{}, plain:{})".format(err1, str(e2)))


# ── IMAP 폴더 자동 탐지 헬퍼 ─────────────────────────
# 다우오피스/IMAP 서버마다 폴더명이 다름.
# - 보낸메일함: Sent, Sent Items, Sent Messages, "보낸메일함", "보낸 편지함" 등
# - LIST 명령으로 폴더 목록을 받고, 표준 \\Sent 플래그 또는 이름으로 매칭
def _decode_mailbox_name(raw):
    """IMAP LIST 응답에서 폴더명 부분 추출 (modified UTF-7 디코딩 포함)"""
    try:
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="ignore")
        # LIST 응답 형식: (\HasNoChildren \Sent) "/" "Sent"
        # 마지막 따옴표 안의 값이 폴더명
        if '"' in raw:
            # 마지막 따옴표 쌍
            parts = raw.rsplit('"', 2)
            if len(parts) >= 3:
                return parts[1]
        # 따옴표 없으면 마지막 공백 뒤
        return raw.split()[-1] if raw.split() else ""
    except:
        return ""

def _imap_utf7_decode(name):
    """IMAP modified UTF-7 → UTF-8 (보낸메일함 같은 한글 폴더명용)"""
    if not name or "&" not in name:
        return name
    try:
        # imap_utf7 모듈 없으면 직접 변환
        import re as _re
        def _decode_part(m):
            b64 = m.group(1).replace(",", "/")
            # 패딩
            b64 += "=" * (-len(b64) % 4)
            try:
                return base64.b64decode(b64).decode("utf-16-be")
            except:
                return m.group(0)
        return _re.sub(r"&([A-Za-z0-9+/,]*)-", lambda m: "&" if m.group(1)=="" else _decode_part(m), name)
    except:
        return name

def _imap_utf7_encode(name):
    """UTF-8 → IMAP modified UTF-7 (한글 폴더 select 할 때 필요)"""
    try:
        # ASCII 문자만 있으면 인코딩 불필요
        name.encode("ascii")
        return name
    except UnicodeEncodeError:
        pass
    # 비ASCII 부분만 modified UTF-7로
    import re as _re
    result = []
    buf = []
    for ch in name:
        if ord(ch) < 0x20 or ord(ch) > 0x7e or ch == "&":
            buf.append(ch)
        else:
            if buf:
                # 비ASCII 버퍼 flush
                b64 = base64.b64encode("".join(buf).encode("utf-16-be")).decode("ascii")
                b64 = b64.rstrip("=").replace("/", ",")
                result.append("&" + b64 + "-")
                buf = []
            if ch == "&":
                result.append("&-")
            else:
                result.append(ch)
    if buf:
        b64 = base64.b64encode("".join(buf).encode("utf-16-be")).decode("ascii")
        b64 = b64.rstrip("=").replace("/", ",")
        result.append("&" + b64 + "-")
    return "".join(result)

def _list_folders(mail):
    """IMAP 서버의 폴더 목록 조회 → [(이름, 플래그)]"""
    try:
        typ, data = mail.list()
        if typ != "OK":
            return []
        folders = []
        for raw in data:
            if not raw:
                continue
            try:
                s = raw.decode("utf-8", errors="ignore") if isinstance(raw, bytes) else str(raw)
            except:
                continue
            # 형식: (\HasNoChildren \Sent) "/" "Sent"
            import re as _re
            m = _re.match(r'\((.*?)\)\s+("[^"]*"|\S+)\s+("[^"]+"|\S+)$', s)
            if m:
                flags = m.group(1)
                name_raw = m.group(3).strip('"')
                name = _imap_utf7_decode(name_raw)
                folders.append((name, name_raw, flags))
        return folders
    except Exception as e:
        print("[IMAP] list 오류: {}".format(e))
        return []

def _pick_folder(mail, target):
    """target: 'inbox' 또는 'sent'. 매칭되는 폴더의 raw 이름 반환"""
    if target == "inbox":
        return "INBOX"
    folders = _list_folders(mail)
    # 보낸메일함 후보 (플래그 우선 → 영문명 → 한글명)
    if target == "sent":
        # 1. \Sent 플래그
        for name, raw, flags in folders:
            if "\\Sent" in flags or "\\sent" in flags.lower():
                return raw
        # 2. 영문 표준명
        en_candidates = ["Sent", "Sent Items", "Sent Messages", "Sent Mail", "INBOX.Sent", "INBOX/Sent"]
        for name, raw, flags in folders:
            if name in en_candidates or raw in en_candidates:
                return raw
        # 3. 한글
        kr_candidates = ["보낸메일함", "보낸 메일함", "보낸편지함", "보낸 편지함"]
        for name, raw, flags in folders:
            if name in kr_candidates:
                return raw
        # 4. 이름에 sent/보낸 포함
        for name, raw, flags in folders:
            low = name.lower()
            if "sent" in low or "보낸" in name:
                return raw
    return None


def imap_get_mails(user, password, email_addr, server=None, limit=200, folder="inbox"):
    """팀원 메일 조회. folder: 'inbox' 또는 'sent'"""
    try:
        mail = imap_connect(user, password, server)
        picked = _pick_folder(mail, folder)
        if not picked:
            mail.logout()
            return {"mails": [], "count": 0, "error": "{} 폴더를 찾을 수 없습니다".format(folder)}
        # 한글 폴더명은 modified UTF-7로 인코딩한 raw가 이미 들어있음 (folders[][1])
        # 폴더명을 따옴표로 감싸서 select
        typ, _ = mail.select('"{}"'.format(picked) if " " in picked else picked)
        if typ != "OK":
            mail.logout()
            return {"mails": [], "count": 0, "error": "select 실패: {}".format(picked)}

        _, data = mail.search(None, "ALL")
        ids = data[0].split()
        ids = ids[-limit:][::-1]  # 최신순

        results = []
        for uid in ids:
            try:
                _, msg_data = mail.fetch(uid, "(RFC822.HEADER)")
                msg = email.message_from_bytes(msg_data[0][1])
                subject  = decode_str(msg.get("Subject", ""))
                from_raw = decode_str(msg.get("From", ""))
                to_raw   = msg.get("To", "") or ""
                cc_raw   = msg.get("Cc", "") or ""
                date_str = msg.get("Date", "")
                ct       = msg.get("Content-Type", "")
                try:
                    d = email.utils.parsedate_to_datetime(date_str)
                    date_fmt = d.strftime("%Y. %m. %d. %p %I:%M").replace("AM","오전").replace("PM","오후")
                    date_raw = d.isoformat()
                except:
                    date_fmt = date_str
                    date_raw = date_str

                results.append({
                    "cacheKey":    f"{email_addr}/imap/{folder}/{uid.decode()}",
                    "uid":         uid.decode(),
                    "owner":       email_addr,
                    "ownerName":   "",
                    "subject":     subject,
                    "from":        from_raw,
                    "to":          parse_addresses(to_raw),
                    "cc":          parse_addresses(cc_raw),
                    "date":        date_fmt,
                    "dateRaw":     date_raw,
                    "hasAttachment": "mixed" in ct.lower(),
                    "attachments": [],
                    "headerLoaded": True,
                    "bodyLoaded":  False,
                    "body":        "",
                    "source":      "imap",
                    "folder":      folder,
                })
            except Exception as e:
                continue

        mail.logout()
        return {"mails": results, "count": len(results), "folder": folder, "folder_raw": picked, "error": None}
    except Exception as e:
        return {"mails": [], "count": 0, "error": str(e)}


def _try_one_fetch(user, password, server, folder, uid_str, cmd_kind):
    """한 번의 시도: 새 연결 만들고 한 명령만 던짐. 성공시 bytes, 실패시 raise"""
    mail = imap_connect(user, password, server)
    try:
        picked = _pick_folder(mail, folder)
        if not picked:
            raise Exception("폴더 없음: " + folder)
        mail.select('"{}"'.format(picked) if (" " in picked or any(ord(c)>127 for c in picked)) else picked)
        uid_b = uid_str.encode()
        if cmd_kind == "uid_rfc822":
            typ, data = mail.uid("FETCH", uid_b, "(RFC822)")
        elif cmd_kind == "seq_rfc822":
            typ, data = mail.fetch(uid_b, "(RFC822)")
        elif cmd_kind == "uid_peek":
            typ, data = mail.uid("FETCH", uid_b, "(BODY.PEEK[])")
        elif cmd_kind == "uid_body":
            typ, data = mail.uid("FETCH", uid_b, "(BODY[])")
        elif cmd_kind == "seq_peek":
            typ, data = mail.fetch(uid_b, "(BODY.PEEK[])")
        else:
            raise Exception("unknown cmd: " + cmd_kind)
        if typ != "OK":
            raise Exception("typ={} data={}".format(typ, str(data)[:80]))
        for item in data:
            if isinstance(item, tuple) and len(item) >= 2:
                return item[1]
        raise Exception("응답에 메시지 본문 없음")
    finally:
        try: mail.logout()
        except: pass


def _imap_fetch_raw_body_safe(user, password, server, folder, uid_str):
    """
    각 시도마다 새 IMAP 연결 만들어 첫 성공하는 명령 사용.
    다우오피스가 응답 파싱 실패시 연결 상태가 망가지므로 시도마다 재연결 필요.

    ⚠️ SEQ FETCH는 사용하지 않음! (UID는 메일 고유 번호이지만, sequence number는
    메일함 인덱스라 다른 메일을 가리킬 수 있음. 다우오피스에서 seq fallback이
    잘못된 메일 본문을 반환하는 사례 확인됨)
    """
    attempts = [
        ("uid_rfc822", "UID FETCH RFC822"),
        ("uid_peek",   "UID FETCH BODY.PEEK[]"),
        ("uid_body",   "UID FETCH BODY[]"),
        # SEQ 기반은 제거 - 다른 메일을 fetch할 위험
    ]
    errors = []
    for kind, label in attempts:
        try:
            print("[FETCH] {} 시도: {}/{} uid={}".format(label, folder, "?", uid_str))
            raw = _try_one_fetch(user, password, server, folder, uid_str, kind)
            if raw:
                print("[FETCH] ✓ {} 성공 ({} bytes)".format(label, len(raw)))
                return raw
        except Exception as e:
            msg = str(e)[:150]
            errors.append("{}: {}".format(label, msg))
            print("[FETCH] ✗ {} 실패: {}".format(label, msg))
            continue
    raise Exception("모든 UID FETCH 실패\n" + "\n".join(errors))


def imap_get_body(user, password, uid, server=None, folder="inbox"):
    """특정 메일 본문 가져오기. folder: 'inbox' 또는 'sent'
    다우오피스 응답 비표준 케이스 대비 5단계 fallback."""
    try:
        raw_bytes = _imap_fetch_raw_body_safe(user, password, server, folder, uid)
        if not raw_bytes:
            return {"body":"", "attachments":[], "hasAttachment":False, "error":"빈 응답"}

        try:
            msg = email.message_from_bytes(raw_bytes)
        except Exception as parse_err:
            return {"body":"", "attachments":[], "hasAttachment":False,
                    "error": "EML 파싱 오류: " + str(parse_err)[:200]}

        body = extract_body(msg)
        attachments = extract_attachments(msg)
        return {
            "body": body,
            "attachments": [{"name": a["name"], "content_type": a.get("content_type","")} for a in attachments],
            "hasAttachment": len(attachments) > 0,
            "error": None
        }
    except Exception as e:
        return {"body":"", "attachments":[], "hasAttachment":False, "error": str(e)[:500]}


@app.route("/api/imap/mails", methods=["POST"])
def api_imap_mails():
    """팀원별 IMAP 메일 목록 조회
    Body: { user, pass, email, server(optional), limit(optional), folder(optional: "inbox"|"sent") }
    """
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "JSON 데이터 없음"}), 400

        user     = data.get("user", "")
        password = data.get("pass", "")
        email_addr = data.get("email", "")
        server   = data.get("server", None)
        limit    = int(data.get("limit", 200))
        folder   = data.get("folder", "inbox")  # ★ inbox 또는 sent

        if not user or not password:
            return jsonify({"error": "user/pass 필요"}), 400

        result = imap_get_mails(user, password, email_addr, server, limit, folder)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/imap/body", methods=["POST"])
def api_imap_body():
    """특정 메일 본문 조회
    Body: { user, pass, uid, server(optional), folder(optional: "inbox"|"sent") }
    """
    try:
        data = request.get_json()
        user     = data.get("user", "")
        password = data.get("pass", "")
        uid      = data.get("uid", "")
        server   = data.get("server", None)
        folder   = data.get("folder", "inbox")  # ★

        if not user or not password or not uid:
            return jsonify({"error": "user/pass/uid 필요"}), 400

        result = imap_get_body(user, password, uid, server, folder)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── 디버그: 다우오피스 폴더 목록 확인용 ─────────────
@app.route("/api/imap/folders", methods=["POST"])
def api_imap_folders():
    """팀원의 IMAP 폴더 목록 (디버그/탐색용)
    Body: { user, pass, server(optional) }
    응답: [{"name":..., "raw":..., "flags":...}], inbox_detected, sent_detected
    """
    try:
        data = request.get_json() or {}
        user     = data.get("user", "")
        password = data.get("pass", "")
        server   = data.get("server", None)
        if not user or not password:
            return jsonify({"error": "user/pass 필요"}), 400

        mail = imap_connect(user, password, server)
        folders_raw = _list_folders(mail)
        folders = [{"name": n, "raw": r, "flags": f} for n, r, f in folders_raw]
        inbox = _pick_folder(mail, "inbox")
        sent  = _pick_folder(mail, "sent")
        mail.logout()
        return jsonify({
            "folders": folders,
            "inbox_picked": inbox,
            "sent_picked": sent,
            "total": len(folders)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/imap/test", methods=["POST"])
def api_imap_test():
    try:
        data = request.get_json(force=True)
        user     = data.get("user", "")
        password = data.get("pass", "")
        server   = data.get("server", IMAP_SERVER)

        # 디버그: 받은 값 확인
        print(f"[IMAP TEST] user={user}, pass_len={len(password)}, pass_ascii={password.isascii()}")

        mail = imap_connect(user, password, server)
        _, data2 = mail.select("INBOX")
        count = int(data2[0])
        mail.logout()
        return jsonify({"status": "ok", "count": count, "server": server})
    except Exception as e:
        print(f"[IMAP TEST ERROR] {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 400



# ── IMAP → SP 백업 (비활성화) ──────────────────────
# 인덱스 방식 도입 후 더 이상 필요 없음.
# 새 백업은 PC에서 build_index.py 실행하는 방식으로 처리.
# 이 엔드포인트가 살아있으면 worker가 imap fetch 중 죽어서 서버 전체 영향.
@app.route("/api/imap/backup", methods=["POST"])
def api_imap_backup():
    """비활성화됨. 백업은 build_index.py로 처리."""
    return jsonify({
        "saved": 0,
        "total": 0,
        "errors": [],
        "disabled": True,
        "message": "이 엔드포인트는 비활성화되었습니다. 백업은 build_index.py를 사용하세요."
    })


# 옛날 backup 함수 (참고용, 사용 안 함)
def _legacy_imap_backup_DISABLED():
    try:
        data = request.get_json()
        user       = data.get("user", "")
        password   = data.get("pass", "")
        email_addr = data.get("email", "")
        server     = data.get("server", None)
        sp_token   = data.get("sp_token", "")
        drive_id   = data.get("drive_id", "")
        folder     = data.get("folder", "")
        existing   = set(data.get("existing_keys", []))
        limit      = int(data.get("limit", 500))

        if not all([user, password, sp_token, drive_id, folder]):
            return jsonify({"error": "필수 파라미터 없음"}), 400

        # IMAP 연결
        srv = server or IMAP_SERVER
        mail = imaplib.IMAP4_SSL(srv, 993)
        mail.login(user, password)
        mail.select("INBOX")
        _, data2 = mail.search(None, "ALL")
        ids = data2[0].split()
        ids = ids[-limit:]  # 최신 N통

        saved = 0
        errors = []

        for uid in ids:
            filename = f"{uid.decode().zfill(8)}.eml"
            if filename in existing:
                continue  # 이미 백업됨

            try:
                # eml 전체 다운로드
                _, msg_data = mail.fetch(uid, "(RFC822)")
                eml_bytes = msg_data[0][1]

                # SP에 업로드
                sp_path = f"{folder}/{filename}"
                parts = "/".join(urllib.parse.quote(p, safe="") for p in sp_path.split("/"))
                url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parts}:/content"

                req = urllib.request.Request(
                    url,
                    data=eml_bytes,
                    headers={
                        "Authorization": f"Bearer {sp_token}",
                        "Content-Type": "message/rfc822",
                    },
                    method="PUT"
                )
                with urllib.request.urlopen(req, timeout=30) as r:
                    if r.status in (200, 201):
                        saved += 1
            except Exception as e:
                errors.append(f"{uid.decode()}: {str(e)}")
                continue

        mail.logout()
        return jsonify({
            "saved": saved,
            "total": len(ids),
            "errors": errors[:10],  # 최대 10개만
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/imap/backup/status", methods=["POST"])
def api_imap_backup_status():
    """백업 상태 확인 (IMAP 총 메일 수 vs SP 저장된 수)"""
    try:
        data = request.get_json()
        user     = data.get("user", "")
        password = data.get("pass", "")
        server   = data.get("server", None)

        srv = server or IMAP_SERVER
        mail = imaplib.IMAP4_SSL(srv, 993)
        mail.login(user, password)
        mail.select("INBOX")
        _, data2 = mail.search(None, "ALL")
        total = len(data2[0].split())
        mail.logout()

        return jsonify({"imap_total": total, "status": "ok"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── CalDAV 캘린더 ─────────────────────────────────────
CALDAV_SERVER = os.environ.get("CALDAV_SERVER", "https://gw.enjet.co.kr")
CALDAV_MEMBERS = [
    {"email":"kwkang@enjet.co.kr",   "name":"강경원",  "user":os.environ.get("CALDAV_USER1","kwkang"),   "password":os.environ.get("CALDAV_PASS1","")},
    {"email":"baekhoon@enjet.co.kr", "name":"성백훈",  "user":os.environ.get("CALDAV_USER2","baekhoon"), "password":os.environ.get("CALDAV_PASS2","")},
]

# 다우오피스 CalDAV 인증은 풀 이메일(user@enjet.co.kr) 형식 요구.
# ID만 들어온 경우 자동으로 @enjet.co.kr 붙여서 보정.
def _normalize_caldav_user(user, email=""):
    if not user:
        return email or ""
    if "@" in user:
        return user
    # ID만 들어왔으면 email의 도메인 붙이거나 기본 도메인 사용
    if email and "@" in email:
        return "{}@{}".format(user, email.split("@", 1)[1])
    return "{}@enjet.co.kr".format(user)

def caldav_req(user, password, url, xml_body=None, method="REPORT", depth="1"):
    import base64
    creds = base64.b64encode("{}:{}".format(user, password).encode("utf-8")).decode("ascii")
    hdrs = {"Authorization": "Basic " + creds, "Depth": depth, "Content-Type": "application/xml; charset=utf-8"}
    data = xml_body.encode("utf-8") if xml_body else None
    req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.read().decode("utf-8", errors="ignore")
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="ignore")[:100]
        except:
            body = ""
        print("CalDAV {} {} | {} | {}".format(e.code, url, e.headers.get("WWW-Authenticate",""), body))
        return None
    except Exception as e:
        print("CalDAV ERR {}: {}".format(url, e))
        return None

def _gp(key, block):
    pat = "^" + key + "[;:][^\r\n]*((?:\r?\n[ \t][^\r\n]*)*)"
    m = re.search(pat, block, re.MULTILINE)
    if not m:
        return ""
    v = re.sub("^" + key + "[^:]*:", "", m.group(0))
    return re.sub("\r?\n[ \t]", "", v).strip()

def _pdt(s):
    if not s:
        return ""
    s = re.sub("^[^:]*:", "", s).strip().rstrip("Z")
    try:
        if len(s) == 8:
            return "{}-{}-{}".format(s[:4], s[4:6], s[6:8])
        return "{}-{}-{}T{}:{}".format(s[:4], s[4:6], s[6:8], s[9:11], s[11:13])
    except:
        return s

def _parse_cal_addr(raw):
    """
    ORGANIZER/ATTENDEE 라인에서 이메일과 CN(이름) 추출.
    예: "ATTENDEE;CN=홍길동;ROLE=REQ-PARTICIPANT:mailto:hong@example.com"
    """
    if not raw:
        return {"email": "", "name": ""}
    # CN 파라미터
    cn_m = re.search(r"CN=([^;:]+)", raw, re.IGNORECASE)
    cn = (cn_m.group(1).strip().strip('"') if cn_m else "")
    # mailto: 부분
    em_m = re.search(r"(?:mailto:|MAILTO:)([^;:\s\r\n]+)", raw)
    email_v = (em_m.group(1).strip() if em_m else "")
    if not email_v:
        # 마지막 ':' 이후 값에서 이메일 모양 시도
        tail = raw.rsplit(":", 1)[-1].strip()
        if "@" in tail:
            email_v = tail
    return {"email": email_v, "name": cn}


def _gp_all(key, block):
    """같은 키가 여러 번 나오는 경우 (ATTENDEE) 모두 추출"""
    out = []
    pat = re.compile(r"^" + key + r"([;:][^\r\n]*(?:\r?\n[ \t][^\r\n]*)*)", re.MULTILINE)
    for m in pat.finditer(block):
        v = (key + m.group(1))
        v = re.sub(r"\r?\n[ \t]", "", v).strip()
        out.append(v)
    return out


def parse_ical(text, name, email_addr):
    events = []
    if not text:
        return events
    for block in re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", text, re.DOTALL):
        try:
            raw_start = _gp("DTSTART", block)
            clean_start = re.sub("^[^:]*:", "", raw_start).strip()
            title = _gp("SUMMARY", block).replace("\\n", "\n").replace("\\,", ",")
            if not title:
                continue

            # ORGANIZER
            org_raw = _gp("ORGANIZER", block)
            organizer = _parse_cal_addr("ORGANIZER" + (org_raw if org_raw.startswith((":",";")) else ":"+org_raw)) if org_raw else {"email":"","name":""}

            # ATTENDEE (여러 명)
            attendees = []
            for line in _gp_all("ATTENDEE", block):
                p = _parse_cal_addr(line)
                if p["email"] or p["name"]:
                    attendees.append(p)

            events.append({
                "uid":      _gp("UID", block),
                "title":    title,
                "start":    _pdt(raw_start),
                "end":      _pdt(_gp("DTEND", block)),
                "location": _gp("LOCATION", block).replace("\\,", ","),
                "desc":     _gp("DESCRIPTION", block)[:200].replace("\\n","\n").replace("\\,",","),
                "allDay":   len(clean_start) == 8 or "VALUE=DATE" in raw_start,
                "owner":    email_addr,
                "ownerName": name,
                "organizer": organizer,
                "attendees": attendees,
            })
        except Exception as ex:
            print("iCal err: {}".format(ex))
    return events

def fetch_caldav_events(member, start, end):
    raw_user = member.get("user", "")
    ea   = member.get("email", "")
    user = _normalize_caldav_user(raw_user, ea)  # ★ 풀 이메일로 자동 보정
    pwd  = member["password"]
    srv  = CALDAV_SERVER
    if raw_user != user:
        print("[CAL] {} user 보정: '{}' -> '{}'".format(member.get("name","?"), raw_user, user))
    xml  = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:d="DAV:">'
        "<d:prop><d:getetag/><c:calendar-data/></d:prop>"
        "<c:filter><c:comp-filter name=\"VCALENDAR\"><c:comp-filter name=\"VEVENT\">"
        "<c:time-range start=\"{}T000000Z\" end=\"{}T235959Z\"/>".format(start, end) +
        "</c:comp-filter></c:comp-filter></c:filter>"
        "</c:calendar-query>"
    )
    base = "{}/principals/users/{}/calendars".format(srv, ea)
    urls = [
        base + "/%EB%82%B4%20%EC%9D%BC%EC%A0%95/",          # 내 일정
        base + "/%EB%82%B4%20%EC%9D%BC%EC%A0%95(%EA%B8%B0%EB%B3%B8)/",  # 내 일정(기본)
        base + "/",
        base + "/%EB%82%B4%EC%9D%BC%EC%A0%95/",              # 내일정
        base + "/%EA%B8%B0%EB%B3%B8/",                       # 기본
    ]
    for url in urls:
        print("[CAL] {} -> {}".format(member["name"], url))
        resp = caldav_req(user, pwd, url, xml, "REPORT", "1")
        if resp and "VEVENT" in resp:
            evs = []
            for blk in re.findall(r"<.*?calendar-data[^>]*>(.*?)</.*?calendar-data>", resp, re.DOTALL):
                blk = blk.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
                evs.extend(parse_ical(blk, member["name"], ea))
            print("[CAL] {} OK {}ev".format(member["name"], len(evs)))
            return evs
        elif resp:
            print("[CAL] {} resp {}b no VEVENT".format(member["name"], len(resp)))
    print("[CAL] {} FAIL".format(member["name"]))
    return []

@app.route("/api/calendar", methods=["GET", "POST"])
def api_calendar():
    try:
        today = datetime.now()
        start = (today - timedelta(days=30)).strftime("%Y%m%d")
        end   = (today + timedelta(days=60)).strftime("%Y%m%d")
        mlist = []
        if request.method == "POST":
            try:
                body = request.get_json(force=True) or {}
                start = body.get("start", start)
                end   = body.get("end", end)
                for m in body.get("members", []):
                    if m.get("pass"):
                        mlist.append({"email": m["email"], "name": m.get("name",""), "user": m.get("user",""), "password": m["pass"]})
            except Exception as e:
                print("[CAL] body err: {}".format(e))
        if not mlist:
            mlist = [m for m in CALDAV_MEMBERS if m.get("password")]
        all_events = []
        errors = []
        for m in mlist:
            try:
                all_events.extend(fetch_caldav_events(m, start, end))
            except Exception as e:
                errors.append("{}: {}".format(m.get("name","?"), e))
                print("[CAL ERR] {}: {}".format(m.get("name"), e))

        # ── 중복 제거 ──────────────────────────────
        # 같은 일정이 여러 명의 캘린더에 공유되어 있을 때 dedupe.
        # 우선순위: UID 매칭 → (제목+시작시간+종료시간) 매칭.
        # 병합 시 ownerName 들을 합쳐서 "공유: A, B" 형태로 표시.
        dedup = {}
        order = []
        for ev in all_events:
            uid = (ev.get("uid") or "").strip()
            if uid:
                key = "uid:" + uid
            else:
                key = "ts:{}|{}|{}".format(
                    (ev.get("title") or "").strip(),
                    ev.get("start") or "",
                    ev.get("end") or ""
                )
            if key in dedup:
                # 기존 이벤트에 공유자 정보 추가
                existing = dedup[key]
                shared = existing.setdefault("sharedWith", [])
                this_owner = ev.get("ownerName") or ev.get("owner") or ""
                if this_owner and this_owner != existing.get("ownerName") and this_owner not in shared:
                    shared.append(this_owner)
                # attendees가 더 풍부한 쪽으로 갱신
                if len(ev.get("attendees") or []) > len(existing.get("attendees") or []):
                    existing["attendees"] = ev.get("attendees") or []
                if ev.get("organizer", {}).get("email") and not existing.get("organizer", {}).get("email"):
                    existing["organizer"] = ev["organizer"]
            else:
                dedup[key] = dict(ev)
                order.append(key)

        all_events = [dedup[k] for k in order]
        all_events.sort(key=lambda e: e.get("start", ""))
        return jsonify({"events": all_events, "count": len(all_events), "errors": errors})
    except Exception as e:
        import traceback
        print("[CAL FATAL] " + traceback.format_exc())
        return jsonify({"events": [], "count": 0, "errors": [str(e)]}), 200

@app.route("/api/calendar/test")
def api_calendar_test():
    results = []
    for m in CALDAV_MEMBERS:
        if not m.get("password"):
            results.append({"name": m["name"], "status": "PASS 미설정"})
            continue
        url = "{}/principals/users/{}/".format(CALDAV_SERVER, m["email"])
        user_norm = _normalize_caldav_user(m.get("user",""), m.get("email",""))
        resp = caldav_req(user_norm, m["password"], url,
            '<?xml version="1.0"?><propfind xmlns="DAV:"><prop><current-user-principal/></prop></propfind>',
            "PROPFIND", "0")
        results.append({"name": m["name"], "status": "성공" if resp else "실패"})
    return jsonify(results)


# ── 진단용: 사용자별 CalDAV 다양한 메서드/헤더 조합 테스트 ─────────
@app.route("/api/calendar/diagnose", methods=["POST"])
def api_calendar_diagnose():
    """
    POST body: {"user":"hdlee", "email":"hdlee@enjet.co.kr", "password":"..."}
    여러 메서드/헤더 조합으로 어떤 요청은 통과하고 어떤 건 막히는지 확인
    """
    import base64
    body = request.get_json(force=True) or {}
    user = body.get("user", "")
    email = body.get("email", "")
    password = body.get("password", "")
    if not (user and email and password):
        return jsonify({"error": "user, email, password 필요"}), 400

    srv = CALDAV_SERVER
    cal_url = "{}/principals/users/{}/calendars/".format(srv, email)
    inner_url = "{}/principals/users/{}/calendars/%EB%82%B4%20%EC%9D%BC%EC%A0%95/".format(srv, email)
    creds = base64.b64encode("{}:{}".format(user, password).encode("utf-8")).decode("ascii")

    UA_BROWSER = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

    tests = [
        # 1. 가장 단순 GET (브라우저 흉내) - calendars/
        {"name": "GET calendars/ +UA", "method": "GET", "url": cal_url, "depth": None, "ua": UA_BROWSER, "body": None, "ctype": None},
        # 2. GET 내 일정/
        {"name": "GET 내 일정/ +UA", "method": "GET", "url": inner_url, "depth": None, "ua": UA_BROWSER, "body": None, "ctype": None},
        # 3. PROPFIND calendars/ + UA
        {"name": "PROPFIND calendars/ +UA", "method": "PROPFIND", "url": cal_url, "depth": "1", "ua": UA_BROWSER,
         "body": '<?xml version="1.0"?><propfind xmlns="DAV:"><prop><displayname/></prop></propfind>',
         "ctype": "application/xml; charset=utf-8"},
        # 4. PROPFIND calendars/ NO UA (현재 코드 패턴)
        {"name": "PROPFIND calendars/ NO-UA", "method": "PROPFIND", "url": cal_url, "depth": "1", "ua": None,
         "body": '<?xml version="1.0"?><propfind xmlns="DAV:"><prop><displayname/></prop></propfind>',
         "ctype": "application/xml; charset=utf-8"},
        # 5. REPORT 내 일정/ + UA (실제 사용 메서드)
        {"name": "REPORT 내 일정/ +UA", "method": "REPORT", "url": inner_url, "depth": "1", "ua": UA_BROWSER,
         "body": '<?xml version="1.0" encoding="utf-8"?><c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:d="DAV:"><d:prop><d:getetag/><c:calendar-data/></d:prop><c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT"/></c:comp-filter></c:filter></c:calendar-query>',
         "ctype": "application/xml; charset=utf-8"},
        # 6. REPORT 내 일정/ NO UA (현재 코드와 동일)
        {"name": "REPORT 내 일정/ NO-UA", "method": "REPORT", "url": inner_url, "depth": "1", "ua": None,
         "body": '<?xml version="1.0" encoding="utf-8"?><c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:d="DAV:"><d:prop><d:getetag/><c:calendar-data/></d:prop><c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT"/></c:comp-filter></c:filter></c:calendar-query>',
         "ctype": "application/xml; charset=utf-8"},
    ]

    results = []
    for t in tests:
        hdrs = {"Authorization": "Basic " + creds}
        if t["depth"] is not None:
            hdrs["Depth"] = t["depth"]
        if t["ua"]:
            hdrs["User-Agent"] = t["ua"]
        if t["ctype"]:
            hdrs["Content-Type"] = t["ctype"]
        data = t["body"].encode("utf-8") if t["body"] else None
        try:
            req = urllib.request.Request(t["url"], data=data, headers=hdrs, method=t["method"])
            with urllib.request.urlopen(req, timeout=20) as r:
                txt = r.read().decode("utf-8", errors="ignore")
                results.append({
                    "test": t["name"],
                    "status": r.status,
                    "len": len(txt),
                    "preview": txt[:200],
                    "result": "✅"
                })
        except urllib.error.HTTPError as e:
            try:
                err_body = e.read().decode("utf-8", errors="ignore")[:200]
            except:
                err_body = ""
            results.append({
                "test": t["name"],
                "status": e.code,
                "www_auth": e.headers.get("WWW-Authenticate", ""),
                "preview": err_body,
                "result": "❌"
            })
        except Exception as e:
            results.append({"test": t["name"], "error": str(e), "result": "💥"})

    return jsonify({
        "user": user,
        "email": email,
        "calendars_url": cal_url,
        "inner_url": inner_url,
        "results": results
    })


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📋 ENJET 전자결재 시스템 - Teams 연동 백엔드 (Phase 2)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
import hmac
import hashlib
import time
from urllib.parse import quote

# ──────────────────────────────────────────────────────────────
# 환경 변수
# ──────────────────────────────────────────────────────────────
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
SP_SITE_PATH = os.environ.get("SP_SITE_PATH", "enjetco.sharepoint.com:/sites/06.BackOffice")
APPROVAL_MASTER_PATH = os.environ.get("APPROVAL_MASTER_PATH", "📋 전자결재/99.Dashboard(real-time)/approval-master.json")
HR_FILE_PATH = os.environ.get("HR_FILE_PATH", "👥 인사·총무·복지/01.인사/[인사현황] 20260507_인사관리.xlsx")
HMAC_SECRET = os.environ.get("HMAC_SECRET", "enjet-default-secret-change-me")
FRONTEND_URL = os.environ.get("FRONTEND_URL", "https://bhseong77.github.io/enjet-approval")

# 토큰/사이트 캐시
_app_token_cache = {"token": None, "expires": 0}
_sp_ids_cache = {"site_id": None, "drive_id": None}


# ──────────────────────────────────────────────────────────────
# Microsoft Graph 인증 (Client Credentials Flow)
# ──────────────────────────────────────────────────────────────
def get_app_token():
    """Application 토큰 발급 (1시간 캐시)"""
    now_ts = time.time()
    if _app_token_cache["token"] and _app_token_cache["expires"] > now_ts + 60:
        return _app_token_cache["token"]
    
    if not AZURE_CLIENT_ID or not AZURE_CLIENT_SECRET:
        raise Exception("AZURE_CLIENT_ID 또는 AZURE_CLIENT_SECRET이 설정되지 않음")
    
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }).encode()
    
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    
    with urllib.request.urlopen(req, timeout=30) as r:
        result = json.loads(r.read().decode())
    
    _app_token_cache["token"] = result["access_token"]
    _app_token_cache["expires"] = now_ts + result.get("expires_in", 3600)
    return _app_token_cache["token"]


def graph_request(method, path, body=None, token=None):
    """Microsoft Graph API 호출 헬퍼"""
    if token is None:
        token = get_app_token()
    
    url = f"https://graph.microsoft.com/v1.0{path}"
    data = None
    headers = {"Authorization": f"Bearer {token}"}
    
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            content = r.read()
            if not content:
                return None
            try:
                return json.loads(content.decode("utf-8"))
            except:
                return content
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        raise Exception(f"Graph {e.code}: {error_body[:300]}")


# ──────────────────────────────────────────────────────────────
# SharePoint 헬퍼
# ──────────────────────────────────────────────────────────────
def get_sp_ids():
    """SP 사이트/드라이브 ID 캐싱"""
    if _sp_ids_cache["site_id"] and _sp_ids_cache["drive_id"]:
        return _sp_ids_cache["site_id"], _sp_ids_cache["drive_id"]
    
    # 사이트 ID
    site_data = graph_request("GET", f"/sites/{SP_SITE_PATH}")
    _sp_ids_cache["site_id"] = site_data["id"]
    
    # 드라이브 ID
    drives = graph_request("GET", f"/sites/{_sp_ids_cache['site_id']}/drives")
    drive = next((d for d in drives["value"] if d["name"] in ["Documents", "문서"]), drives["value"][0])
    _sp_ids_cache["drive_id"] = drive["id"]
    
    return _sp_ids_cache["site_id"], _sp_ids_cache["drive_id"]


def sp_read_json(path):
    """SharePoint JSON 파일 읽기"""
    _, drive_id = get_sp_ids()
    encoded_path = quote(path, safe="")
    try:
        result = graph_request("GET", f"/drives/{drive_id}/root:/{encoded_path}:/content")
        return result
    except Exception as e:
        if "404" in str(e):
            return None
        raise


def sp_write_json(path, data):
    """SharePoint JSON 파일 쓰기"""
    _, drive_id = get_sp_ids()
    encoded_path = quote(path, safe="")
    return graph_request("PUT", f"/drives/{drive_id}/root:/{encoded_path}:/content", body=data)


# ──────────────────────────────────────────────────────────────
# HMAC 서명 (카드 변조 방지)
# ──────────────────────────────────────────────────────────────
def make_signature(approval_id):
    """결재 ID 기반 HMAC 서명 생성"""
    return hmac.new(
        HMAC_SECRET.encode(),
        approval_id.encode(),
        hashlib.sha256
    ).hexdigest()[:16]


def verify_signature(approval_id, sig):
    """서명 검증"""
    expected = make_signature(approval_id)
    return hmac.compare_digest(expected, sig or "")


# ──────────────────────────────────────────────────────────────
# Adaptive Card 생성
# ──────────────────────────────────────────────────────────────
def build_form_facts(approval):
    """양식별 주요 정보를 FactSet으로 변환"""
    form_id = approval.get("formId", "")
    data = approval.get("formData", {})
    facts = []
    
    if form_id == "VACATION":
        facts.append({"title": "휴가종류", "value": data.get("vacationType", "-")})
        period = f"{data.get('startDate', '-')} ~ {data.get('endDate', '-')}"
        if data.get("vacationType") == "반차":
            period += f" ({data.get('halfStart', '')} ~ {data.get('halfEnd', '')})"
        facts.append({"title": "기간", "value": period})
        facts.append({"title": "일수", "value": f"{data.get('days', '-')}일"})
        if data.get("reason"):
            facts.append({"title": "사유", "value": data["reason"][:100]})
    
    elif form_id == "EXPENSE":
        facts.append({"title": "지출구분", "value": data.get("expenseType", "-")})
        try:
            amt = int(float(data.get("amount", 0)))
            facts.append({"title": "금액", "value": f"{amt:,}원"})
        except:
            pass
        facts.append({"title": "거래처", "value": data.get("vendor", "-")})
        facts.append({"title": "결제방법", "value": data.get("payMethod", "-")})
    
    elif form_id == "BIZ_TRIP":
        facts.append({"title": "출장구분", "value": data.get("tripType", "-")})
        facts.append({"title": "출장지", "value": data.get("tripLocation", "-")})
        period = f"{data.get('tripStart', '-')} ~ {data.get('tripEnd', '-')}"
        facts.append({"title": "기간", "value": f"{period} ({data.get('tripDays', '-')}일)"})
        try:
            total = int(float(data.get("costTotal", 0)))
            if total > 0:
                facts.append({"title": "예상비용", "value": f"{total:,}원"})
        except:
            pass
    
    elif form_id == "PURCHASE_REQUEST":
        facts.append({"title": "프로젝트", "value": data.get("pjCode", "-")})
        try:
            amt = int(float(data.get("totalAmount", 0)))
            facts.append({"title": "금액", "value": f"{amt:,}원"})
        except:
            pass
        facts.append({"title": "거래처", "value": data.get("vendor", "-")})
    
    elif form_id == "QUOTE_APPROVAL":
        facts.append({"title": "고객사", "value": data.get("customer", "-")})
        try:
            amt = int(float(data.get("quoteAmount", 0)))
            facts.append({"title": "견적금액", "value": f"{amt:,}원"})
        except:
            pass
        if data.get("discount"):
            facts.append({"title": "할인율", "value": f"{data['discount']}%"})
    
    return facts


def build_approval_card(approval, recipient_email):
    """결재 요청 Adaptive Card 생성"""
    sig = make_signature(approval["id"])
    api_base = request.host_url.rstrip("/") if request else ""
    
    facts = [
        {"title": "기안자", "value": f"{approval['drafter']['name']} {approval['drafter']['rank']}"},
        {"title": "소속", "value": f"{approval['drafter']['dept']} · {approval['drafter'].get('team', '')}".strip(" ·")},
        {"title": "기안일", "value": approval.get("createdAt", "")[:10]},
    ]
    facts.extend(build_form_facts(approval))
    
    cur_step_idx = approval.get("currentStep", 0)
    total_steps = len(approval.get("approvalLine", []))
    progress_text = f"{cur_step_idx + 1} / {total_steps} 단계"
    
    is_urgent = approval.get("urgency") == "긴급"
    title_color = "Attention" if is_urgent else "Accent"
    title_prefix = "🚨 [긴급] " if is_urgent else "📥 "
    
    card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {
                "type": "Container",
                "items": [
                    {"type": "TextBlock", "text": f"{title_prefix}결재 요청 ({progress_text})", "weight": "Bolder", "color": title_color, "size": "Small"},
                    {"type": "TextBlock", "text": f"{approval.get('formIcon', '📄')} {approval.get('formName', '')}", "size": "Large", "weight": "Bolder", "spacing": "Small"},
                    {"type": "TextBlock", "text": approval.get("title", ""), "wrap": True, "size": "Medium", "spacing": "Small"}
                ]
            },
            {"type": "FactSet", "facts": facts},
            {
                "type": "Input.Text",
                "id": "comment",
                "placeholder": "결재 의견 (선택사항)",
                "isMultiline": True,
                "maxLength": 500
            }
        ],
        "actions": [
            {
                "type": "Action.Http",
                "title": "✓ 승인",
                "method": "POST",
                "url": f"{api_base}/api/approval/action",
                "headers": [{"name": "Content-Type", "value": "application/json"}],
                "body": json.dumps({
                    "id": approval["id"],
                    "action": "승인",
                    "approver": recipient_email,
                    "comment": "{{comment.value}}",
                    "sig": sig
                }, ensure_ascii=False),
                "style": "positive"
            },
            {
                "type": "Action.Http",
                "title": "✕ 반려",
                "method": "POST",
                "url": f"{api_base}/api/approval/action",
                "headers": [{"name": "Content-Type", "value": "application/json"}],
                "body": json.dumps({
                    "id": approval["id"],
                    "action": "반려",
                    "approver": recipient_email,
                    "comment": "{{comment.value}}",
                    "sig": sig
                }, ensure_ascii=False),
                "style": "destructive"
            },
            {
                "type": "Action.OpenUrl",
                "title": "📋 상세보기",
                "url": f"{FRONTEND_URL}/?id={approval['id']}"
            }
        ]
    }
    return card


def build_status_card(approval, status_type):
    """결재 결과 알림 카드 (완료/반려/참조)"""
    icon_map = {
        "완료": ("✅", "Good", "결재가 완료되었습니다"),
        "반려": ("❌", "Attention", "결재가 반려되었습니다"),
        "참조": ("📋", "Accent", "참조 결재 알림"),
        "다음": ("🔔", "Accent", "결재가 진행 중입니다")
    }
    icon, color, msg = icon_map.get(status_type, ("📄", "Default", "알림"))
    
    facts = [
        {"title": "기안자", "value": f"{approval['drafter']['name']} {approval['drafter']['rank']}"},
        {"title": "양식", "value": approval.get("formName", "")},
        {"title": "제목", "value": approval.get("title", "")},
    ]
    
    # 마지막 결재 정보
    last_action = None
    for step in approval.get("approvalLine", []):
        if step.get("status") in ["승인", "반려"]:
            last_action = step
    
    if last_action:
        action_text = "승인" if last_action["status"] == "승인" else "반려"
        facts.append({"title": "처리자", "value": f"{last_action['name']} {last_action.get('rank', '')} · {action_text}"})
        if last_action.get("comment"):
            facts.append({"title": "결재 의견", "value": last_action["comment"][:200]})
    
    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": f"{icon} {msg}", "weight": "Bolder", "color": color, "size": "Medium"},
            {"type": "FactSet", "facts": facts}
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "📋 상세보기",
                "url": f"{FRONTEND_URL}/?id={approval['id']}"
            }
        ]
    }


# ──────────────────────────────────────────────────────────────
# Teams DM 발송
# ──────────────────────────────────────────────────────────────
def send_teams_dm(recipient_email, card):
    """1대1 채팅으로 Adaptive Card 발송 (앱 권한 사용)"""
    try:
        # 1. 수신자 user ID 조회
        recipient = graph_request("GET", f"/users/{quote(recipient_email)}")
        recipient_id = recipient["id"]
        
        # 2. 발신자(앱) 식별 - installedApp 사용
        # Application 권한으로 chat 생성 시 onBehalfOf 방식 사용
        # 대안: Microsoft Graph Mail 발송으로 대체 (간단함)
        
        # 채팅 메시지 발송 (앱이 보내는 시스템 메시지)
        # 주의: Chat.ReadWrite.All 권한 필요
        try:
            # 1대1 채팅 생성 시도
            chat_data = graph_request("POST", "/chats", body={
                "chatType": "oneOnOne",
                "members": [
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{recipient_id}')"
                    }
                ]
            })
            chat_id = chat_data["id"]
        except Exception as e:
            # 앱 권한으로 1대1 채팅 생성이 안 되면 메일로 폴백
            print(f"Teams DM 실패, 메일 폴백: {e}")
            return send_email_fallback(recipient_email, card)
        
        # 메시지 발송
        result = graph_request("POST", f"/chats/{chat_id}/messages", body={
            "body": {
                "contentType": "html",
                "content": '<attachment id="1"></attachment>'
            },
            "attachments": [{
                "id": "1",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": json.dumps(card, ensure_ascii=False)
            }]
        })
        return {"success": True, "messageId": result.get("id"), "method": "teams"}
    
    except Exception as e:
        print(f"Teams DM 오류: {e}")
        return send_email_fallback(recipient_email, card)


def send_email_fallback(recipient_email, card):
    """Teams 발송 실패 시 이메일로 폴백"""
    try:
        # 카드 내용을 HTML 이메일로 변환
        title = ""
        facts_html = ""
        for item in card.get("body", []):
            if item.get("type") == "TextBlock" and not title:
                title = item.get("text", "")
            elif item.get("type") == "FactSet":
                for fact in item.get("facts", []):
                    facts_html += f"<tr><td style='padding:6px 12px;background:#f3f4f6;font-weight:600'>{fact['title']}</td><td style='padding:6px 12px'>{fact['value']}</td></tr>"
        
        actions_html = ""
        for action in card.get("actions", []):
            if action.get("type") == "Action.OpenUrl":
                url = action.get("url", "#")
                actions_html += f'<a href="{url}" style="display:inline-block;padding:10px 20px;background:#1a6fd4;color:#fff;text-decoration:none;border-radius:6px;margin:4px">{action["title"]}</a>'
        
        html_body = f"""
        <html><body style="font-family:Arial,sans-serif;max-width:600px;margin:auto">
          <h2 style="color:#1a6fd4">{title}</h2>
          <table style="width:100%;border-collapse:collapse;margin:14px 0">{facts_html}</table>
          <div style="margin-top:20px">{actions_html}</div>
          <p style="font-size:12px;color:#999;margin-top:20px">⚠️ Teams 알림이 전송되지 않아 이메일로 발송되었습니다.</p>
        </body></html>
        """
        
        # Graph API로 메일 발송
        # 발신자: 첫 번째 사용가능한 사용자 (또는 시스템 계정)
        sender_email = recipient_email  # 자기 자신에게 보내기 (delegated 없이 가능한 방법)
        
        # 앱 권한으로 메일 발송 - sendMail
        graph_request("POST", f"/users/{quote(sender_email)}/sendMail", body={
            "message": {
                "subject": title,
                "body": {"contentType": "HTML", "content": html_body},
                "toRecipients": [{"emailAddress": {"address": recipient_email}}]
            },
            "saveToSentItems": "false"
        })
        return {"success": True, "method": "email"}
    except Exception as e:
        print(f"이메일 폴백도 실패: {e}")
        return {"success": False, "error": str(e)}


# ──────────────────────────────────────────────────────────────
# 결재 데이터 처리
# ──────────────────────────────────────────────────────────────
def load_approval_master():
    """결재 마스터 JSON 로드"""
    data = sp_read_json(APPROVAL_MASTER_PATH)
    if data is None:
        return {"approvals": []}
    return data


def save_approval_master(data):
    """결재 마스터 JSON 저장"""
    data["lastUpdate"] = datetime.utcnow().isoformat() + "Z"
    sp_write_json(APPROVAL_MASTER_PATH, data)


# ──────────────────────────────────────────────────────────────
# 결재 API 엔드포인트
# ──────────────────────────────────────────────────────────────
@app.route("/api/approval/health", methods=["GET"])
def approval_health():
    """결재 시스템 헬스 체크"""
    try:
        token_ok = bool(get_app_token())
        site_id, drive_id = get_sp_ids()
        return jsonify({
            "status": "ok",
            "azure_auth": token_ok,
            "site_id": site_id[:30] + "..." if site_id else None,
            "drive_id": drive_id[:30] + "..." if drive_id else None,
            "frontend_url": FRONTEND_URL,
            "master_path": APPROVAL_MASTER_PATH
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/api/approval/notify", methods=["POST"])
def approval_notify():
    """
    결재자에게 Teams DM 발송
    프론트엔드(HTML)가 결재 상신 후 호출
    
    Body: {
        "approvalId": "AP-260522-001",
        "recipientEmail": "approver@enjet.co.kr",
        "notifyType": "approval" | "complete" | "rejected" | "cc"
    }
    """
    try:
        payload = request.get_json(force=True)
        approval_id = payload.get("approvalId")
        recipient = payload.get("recipientEmail")
        notify_type = payload.get("notifyType", "approval")
        
        if not approval_id or not recipient:
            return jsonify({"error": "approvalId와 recipientEmail 필수"}), 400
        
        # 결재 데이터 로드
        master = load_approval_master()
        approval = next((a for a in master["approvals"] if a["id"] == approval_id), None)
        if not approval:
            return jsonify({"error": "결재 문서 없음"}), 404
        
        # 카드 선택
        if notify_type == "approval":
            card = build_approval_card(approval, recipient)
        elif notify_type == "complete":
            card = build_status_card(approval, "완료")
        elif notify_type == "rejected":
            card = build_status_card(approval, "반려")
        elif notify_type == "cc":
            card = build_status_card(approval, "참조")
        else:
            card = build_status_card(approval, "다음")
        
        result = send_teams_dm(recipient, card)
        return jsonify(result)
    
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()[:500]}), 500


@app.route("/api/approval/action", methods=["POST"])
def approval_action():
    """
    Teams Adaptive Card [승인]/[반려] 버튼 클릭 시 호출
    
    Body: {
        "id": "AP-260522-001",
        "action": "승인" | "반려",
        "approver": "user@enjet.co.kr",
        "comment": "결재 의견",
        "sig": "HMAC 서명"
    }
    """
    try:
        payload = request.get_json(force=True)
        approval_id = payload.get("id")
        action = payload.get("action")
        approver = (payload.get("approver") or "").lower()
        comment = payload.get("comment", "").strip()
        sig = payload.get("sig", "")
        
        # 검증 1: HMAC 서명
        if not verify_signature(approval_id, sig):
            return jsonify({"error": "잘못된 서명"}), 403
        
        # 검증 2: 액션 종류
        if action not in ["승인", "반려"]:
            return jsonify({"error": "action은 승인 또는 반려"}), 400
        
        # 결재 데이터 로드
        master = load_approval_master()
        approval_idx = next((i for i, a in enumerate(master["approvals"]) if a["id"] == approval_id), -1)
        if approval_idx < 0:
            return jsonify({"error": "결재 문서 없음"}), 404
        
        approval = master["approvals"][approval_idx]
        
        # 검증 3: 상태
        if approval.get("status") != "진행중":
            return jsonify({"error": f"이미 처리된 결재입니다 (상태: {approval.get('status')})"}), 400
        
        # 검증 4: 현재 결재자 확인
        cur_step_idx = approval.get("currentStep", 0)
        if cur_step_idx >= len(approval.get("approvalLine", [])):
            return jsonify({"error": "결재 단계 오류"}), 400
        
        cur_step = approval["approvalLine"][cur_step_idx]
        if (cur_step.get("email") or "").lower() != approver:
            return jsonify({"error": "결재 권한 없음"}), 403
        if cur_step.get("status") != "대기":
            return jsonify({"error": "이미 처리됨"}), 400
        
        # 결재 처리
        now_iso = datetime.utcnow().isoformat() + "Z"
        cur_step["status"] = action
        cur_step["actionDate"] = now_iso
        cur_step["comment"] = comment
        approval["updatedAt"] = now_iso
        
        next_recipient = None
        notify_type = None
        
        if action == "반려":
            approval["status"] = "반려"
            # 기안자에게 반려 알림
            next_recipient = approval["drafter"]["email"]
            notify_type = "rejected"
        else:  # 승인
            approval["currentStep"] = cur_step_idx + 1
            if approval["currentStep"] >= len(approval["approvalLine"]):
                # 모든 결재 완료
                approval["status"] = "완료"
                approval["completedAt"] = now_iso
                next_recipient = approval["drafter"]["email"]
                notify_type = "complete"
            else:
                # 다음 결재자에게 카드 발송
                next_step = approval["approvalLine"][approval["currentStep"]]
                next_recipient = next_step["email"]
                notify_type = "approval"
        
        # 저장
        master["approvals"][approval_idx] = approval
        save_approval_master(master)
        
        # 다음 알림 발송 (별도 try - 실패해도 결재는 처리됨)
        try:
            if notify_type == "approval":
                card = build_approval_card(approval, next_recipient)
            elif notify_type == "complete":
                card = build_status_card(approval, "완료")
            elif notify_type == "rejected":
                card = build_status_card(approval, "반려")
            send_teams_dm(next_recipient, card)
        except Exception as notify_err:
            print(f"다음 알림 발송 실패: {notify_err}")
        
        # 참조자에게도 알림 (결재 완료 시만)
        if approval["status"] == "완료" and approval.get("ccList"):
            for cc in approval["ccList"]:
                try:
                    cc_card = build_status_card(approval, "참조")
                    send_teams_dm(cc["email"], cc_card)
                except Exception as cc_err:
                    print(f"참조자 알림 실패 ({cc.get('email')}): {cc_err}")
        
        return jsonify({
            "success": True,
            "status": approval["status"],
            "currentStep": approval["currentStep"],
            "message": f"{action} 처리 완료"
        })
    
    except Exception as e:
        import traceback
        print(f"approval_action 오류: {traceback.format_exc()}")
        return jsonify({"error": str(e), "trace": traceback.format_exc()[:500]}), 500


@app.route("/api/approval/status/<approval_id>", methods=["GET"])
def approval_status(approval_id):
    """결재 상태 조회"""
    try:
        master = load_approval_master()
        approval = next((a for a in master["approvals"] if a["id"] == approval_id), None)
        if not approval:
            return jsonify({"error": "결재 문서 없음"}), 404
        return jsonify({
            "id": approval["id"],
            "status": approval["status"],
            "currentStep": approval["currentStep"],
            "totalSteps": len(approval["approvalLine"]),
            "currentApprover": approval["approvalLine"][approval["currentStep"]] if approval["status"] == "진행중" and approval["currentStep"] < len(approval["approvalLine"]) else None
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/approval/resend", methods=["POST"])
def approval_resend():
    """현재 결재자에게 다시 카드 발송 (재알림)"""
    try:
        payload = request.get_json(force=True)
        approval_id = payload.get("approvalId")
        
        master = load_approval_master()
        approval = next((a for a in master["approvals"] if a["id"] == approval_id), None)
        if not approval:
            return jsonify({"error": "결재 문서 없음"}), 404
        if approval["status"] != "진행중":
            return jsonify({"error": "진행중 결재가 아님"}), 400
        
        cur_step = approval["approvalLine"][approval["currentStep"]]
        card = build_approval_card(approval, cur_step["email"])
        result = send_teams_dm(cur_step["email"], card)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 메인
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
