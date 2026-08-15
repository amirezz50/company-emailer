"""Bulk company emailer — mobile-friendly Streamlit app."""

from __future__ import annotations

import time
from io import BytesIO

import pandas as pd
import streamlit as st

from mailer import (
    already_sent,
    clear_sent_state,
    detect_email_column,
    index_to_excel_row,
    is_valid_email,
    load_excel,
    load_state,
    log_result,
    mark_sent,
    render_template,
    sanitize_email,
    send_one,
    slice_rows,
)

st.set_page_config(page_title="إرسال إيميلات الشركات", page_icon="📧", layout="centered")

st.markdown(
    """
    <style>
      .stButton > button { min-height: 3rem; font-size: 1.05rem; }
      textarea { font-size: 1rem !important; }
      .block-container { padding-top: 1.2rem; padding-bottom: 4rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

PROVIDERS = {
    "Gmail": {"host": "smtp.gmail.com", "port": 587, "ssl": False},
    "Brevo (أحسن للوارد من جيميل الشخصي)": {"host": "smtp-relay.brevo.com", "port": 587, "ssl": False},
    "Outlook / Hotmail": {"host": "smtp.office365.com", "port": 587, "ssl": False},
    "Yahoo": {"host": "smtp.mail.yahoo.com", "port": 587, "ssl": False},
    "مخصص": {"host": "", "port": 587, "ssl": False},
}

if "templates" not in st.session_state:
    st.session_state.templates = [
        {
            "name": "التمبلت الأساسي",
            "subject": "{{company}} — استفسار سريع",
            "body": "مرحباً {{name}}،\n\nكتبت لك بخصوص شركة {{company}}. لو الوقت مناسب، أقدر أوضح الفكرة في دقايق.\n\nلو مش مهتم، رد بكلمة إلغاء وأوقف المراسلة.\n\nتحياتي",
        }
    ]
if "sending" not in st.session_state:
    st.session_state.sending = False


def current_template() -> dict:
    names = [t["name"] for t in st.session_state.templates]
    selected = st.session_state.get("active_template", names[0])
    for item in st.session_state.templates:
        if item["name"] == selected:
            return item
    return st.session_state.templates[0]


st.title("إرسال إيميلات للشركات")

st.header("1) حساب الإرسال")
provider = st.selectbox("مزود الإيميل", list(PROVIDERS.keys()))
preset = PROVIDERS[provider]
smtp_host = st.text_input("سيرفر SMTP", value=preset["host"] or "smtp.gmail.com")
conn_mode = st.radio(
    "نوع الاتصال",
    ["STARTTLS — بورت 587 (موصى به لجيميل)", "SSL — بورت 465"],
    index=0,
    horizontal=True,
)
use_ssl = conn_mode.startswith("SSL")
if provider == "مخصص":
    smtp_port = st.number_input(
        "البورت",
        min_value=1,
        max_value=65535,
        value=465 if use_ssl else 587,
    )
else:
    smtp_port = 465 if use_ssl else 587
    st.caption(f"البورت المستخدم: {smtp_port}")

sender_email = st.text_input("الإيميل اللي هيبعت", placeholder="you@gmail.com")
sender_name = st.text_input("الاسم الظاهر للمستقبل (اختياري)", placeholder="شركتك")
smtp_password = st.text_input(
    "كلمة مرور التطبيق / التوكن / App Password",
    type="password",
    help="لـ Gmail استخدم App Password. لو مزود تاني يدعم توكن SMTP حطه هنا.",
)

st.header("2) ملف الإكسيل")
uploaded = st.file_uploader("ارفع ملف Excel", type=["xlsx"])

df = None
if uploaded is not None:
    try:
        df = load_excel(uploaded)
    except Exception as exc:
        st.error(f"مش قادر أقرأ الملف: {exc}")

if df is not None and df.empty:
    st.warning("الشيت فاضي.")
elif df is not None:
    last_excel_row = index_to_excel_row(len(df) - 1)
    st.success(f"اتقرأ {len(df)} صف بيانات. آخر صف في الإكسيل رقم {last_excel_row}.")
    preview = df.copy()
    preview.insert(0, "رقم الصف في Excel", [index_to_excel_row(i) for i in range(len(df))])
    st.dataframe(preview, use_container_width=True, hide_index=True)

    email_guess = detect_email_column(list(df.columns))
    email_col = st.selectbox(
        "عمود الإيميل",
        options=list(df.columns),
        index=list(df.columns).index(email_guess) if email_guess in df.columns else 0,
    )

    st.subheader("نطاق الصفوف")
    st.caption("الأرقام زي ما هي في Excel. الصف 1 عادةً عناوين الأعمدة، وأول بيانات من الصف 2.")
    range_c1, range_c2 = st.columns(2)
    with range_c1:
        start_row = st.number_input("من صف", min_value=2, max_value=last_excel_row, value=2)
    with range_c2:
        end_row = st.number_input(
            "إلى صف",
            min_value=2,
            max_value=last_excel_row,
            value=min(last_excel_row, 50) if last_excel_row >= 2 else 2,
        )

    work = slice_rows(df, int(start_row), int(end_row))
    st.info(f"هيتبعت لـ {len(work)} صف (من {int(start_row)} إلى {int(end_row)}).")
else:
    email_col = None
    work = None
    start_row = 2
    end_row = 2

st.header("3) التمبلت")
st.caption("استخدم اسم العمود بين أقواس: `{{name}}` أو `{{company}}` أو `{{email}}`.")

names = [t["name"] for t in st.session_state.templates]
active = st.selectbox("اختار التمبلت", names, key="active_template")
tpl = current_template()
tpl["subject"] = st.text_input("عنوان الرسالة", value=tpl["subject"], key=f"subj_{active}")
tpl["body"] = st.text_area("نص الرسالة", value=tpl["body"], height=180, key=f"body_{active}")
html_mode = st.checkbox("التمبلت HTML", value=False)

with st.expander("إضافة تمبلت تاني"):
    new_name = st.text_input("اسم التمبلت الجديد", placeholder="عرض سعر")
    if st.button("إضافة") and new_name.strip():
        if new_name.strip() not in names:
            st.session_state.templates.append(
                {"name": new_name.strip(), "subject": "", "body": ""}
            )
            st.rerun()
        else:
            st.warning("الاسم موجود قبل كده.")

if df is not None and work is not None and not work.empty:
    sample = work.iloc[0].to_dict()
    st.subheader("معاينة أول صف")
    st.write("**العنوان:**", render_template(tpl["subject"], sample))
    st.text(render_template(tpl["body"], sample))

st.header("4) الإرسال")
delay = st.slider("ثواني انتظار بين كل إيميل", min_value=2, max_value=60, value=12)
skip_already_sent = st.checkbox(
    "متبعتش تاني لنفس الصف لو اتبعت قبل كده بنجاح",
    value=True,
    help="الصفوف اللي اترسّلها بنجاح تتتخزن محليًا. الصفوف الجديدة أو اللي فشلت تتتبعت.",
)
copy_to_sender = st.checkbox(
    "ابعت نسخة ليا على إيميل المرسل (عشان أتأكد إن الرسالة طلعت)",
    value=True,
)
stop_on_error = st.checkbox("وقف الإرسال لو إيميل فشل (أضمن عشان محدش يتعدى بالغلط)", value=True)
if st.button("امسح سجل الإرسال عشان أعيد المحاولة"):
    clear_sent_state()
    st.success("اتمسح السجل. تقدر تبعت نفس الصفوف تاني.")

ready = bool(
    sender_email
    and smtp_password
    and smtp_host
    and df is not None
    and work is not None
    and not work.empty
    and email_col
)

if st.button("إرسال تجربة لنفسي", disabled=not bool(sender_email and smtp_password and smtp_host)):
    try:
        send_one(
            host=smtp_host,
            port=int(smtp_port),
            username=sender_email,
            password=smtp_password,
            sender_email=sender_email,
            sender_name=sender_name.strip(),
            to_email=sender_email,
            subject="تجربة إرسال",
            body="لو الرسالة دي وصلتك، الإعدادات صحيحة.",
            html=False,
            use_ssl=use_ssl,
            copy_to_sender=False,
        )
        st.success("جوجل قبل التجربة. شوف صندوق **الرسائل المرسلة** في إيميل المرسل، وبعدين الوارد/Spam عندك.")
    except Exception as exc:
        st.error(f"التجربة فشلت: {exc}")

if st.button("إرسال واحد واحد", type="primary", disabled=not ready or st.session_state.sending):
    st.session_state.sending = True

if not ready:
    st.caption("كمّل البيانات فوق: الإيميل، كلمة المرور/التوكن، الملف، والنطاق.")

if st.session_state.sending and work is not None and not work.empty and email_col:
    state = load_state()
    progress = st.progress(0.0, text="بيبدأ الإرسال...")
    status_box = st.empty()
    log_box = st.empty()
    lines: list[str] = []
    sent_count = 0
    fail_count = 0
    skipped_count = 0
    stopped = False
    total = len(work)
    i = 0

    try:
      for i, (idx, row) in enumerate(work.iterrows()):
        excel_row = index_to_excel_row(int(idx))
        to_email = sanitize_email(row[email_col])
        subject = render_template(tpl["subject"], row.to_dict())
        body = render_template(tpl["body"], row.to_dict())
        progress.progress((i) / total, text=f"الصف {excel_row} — {i}/{total}")

        if not is_valid_email(to_email):
            fail_count += 1
            msg = f"صف {excel_row}: إيميل مش صالح ({to_email})"
            lines.append("❌ " + msg)
            log_result(excel_row, to_email, "invalid", subject, msg)
            log_box.write("\n".join(lines))
            if stop_on_error:
                stopped = True
                status_box.error("اتوقف بسبب إيميل مش صالح. صلّح الصف ده وكمل الباقي بعدين.")
                break
            continue

        if skip_already_sent and already_sent(state, excel_row, to_email):
            skipped_count += 1
            lines.append(f"⏭️ صف {excel_row}: اتبعت قبل كده لـ {to_email}")
            log_box.write("\n".join(lines))
            continue

        status_box.info(f"بيبعت الآن للصف {excel_row}: {to_email}")
        try:
            send_one(
                host=smtp_host,
                port=int(smtp_port),
                username=sender_email,
                password=smtp_password,
                sender_email=sender_email,
                sender_name=sender_name.strip(),
                to_email=to_email,
                subject=subject,
                body=body,
                html=html_mode,
                use_ssl=use_ssl,
                copy_to_sender=copy_to_sender,
            )
            mark_sent(state, excel_row, to_email)
            log_result(excel_row, to_email, "sent", subject)
            sent_count += 1
            lines.append(f"✅ صف {excel_row}: اتبعت لـ {to_email}")
        except Exception as exc:
            fail_count += 1
            err = str(exc)
            log_result(excel_row, to_email, "failed", subject, err)
            lines.append(f"❌ صف {excel_row}: فشل لـ {to_email} — {err}")
            log_box.write("\n".join(lines))
            if stop_on_error:
                stopped = True
                status_box.error("اتوقف بعد فشل الإرسال. الصف ده لسه متعملش، وهيكمل من عنده لما تضغط تاني.")
                break

        log_box.write("\n".join(lines))
        if i < total - 1 and not stopped:
            time.sleep(float(delay))
    finally:
        st.session_state.sending = False

    progress.progress(1.0 if not stopped else (i + 1) / total, text="خلص")
    if not stopped and fail_count == 0:
        status_box.success(
            f"تم. جوجل قبل {sent_count} رسالة — واتخطى {skipped_count}. "
            "ده مش معناه إنها وصلت الوارد: افتح إيميل المرسل → الرسائل المرسلة. "
            "وعند المستقبل شوف Spam و All Mail و Promotions."
        )
    elif not stopped:
        status_box.warning(
            f"خلص النطاق. نجح {sent_count} / فشل {fail_count} / متبعتين قبل كده {skipped_count}. "
            "اضغط إرسال تاني عشان يعيد محاولة اللي فشل بس."
        )

    st.download_button(
        "تحميل سجل الإرسال",
        data="\n".join(lines).encode("utf-8"),
        file_name="send_results.txt",
        mime="text/plain",
    )


def make_sample() -> bytes:
    sample = pd.DataFrame(
        {
            "name": ["أحمد علي", "Sara Hassan", "Olivia Taylor"],
            "company": ["شركة النور", "Delta Tech", "Cairo Foods"],
            "email": ["ahmed@example.com", "sara@example.com", "omar@example.com"],
        }
    )
    buf = BytesIO()
    sample.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


st.divider()
st.download_button(
    "تحميل ملف Excel تجريبي",
    data=make_sample(),
    file_name="sample_contacts.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
