# إرسال إيميلات للشركات

برنامج ويب مجاني يقرأ إيميلات الشركات من Excel ويبعت تمبلت واحد (أو أكتر) واحد واحد، مع تحديد الصفوف.

## التشغيل المحلي

```powershell
cd d:\Work\test
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## النشر على Streamlit Cloud (من أي مكان)

1. ارفع المشروع على GitHub.
2. ادخل [share.streamlit.io](https://share.streamlit.io) وسجّل بجيت هاب.
3. New app → اختار المستودع و`app.py`.
4. Settings → Secrets:

```toml
APP_PASSWORD = "كلمة-سر-التطبيق"
```

بعدها هيفتح من الموبايل بأي نت، ومش هيشتغل من غير كلمة السر دي.
