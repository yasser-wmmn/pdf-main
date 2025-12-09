# استخدام صورة بايثون خفيفة
FROM python:3.9-slim

# منع بايثون من كتابة ملفات pyc وتفعيل التخزين المؤقت للإخراج
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# تحديث النظام وتثبيت مكتبات النظام الضرورية (Tesseract و Poppler)
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-ara \
    tesseract-ocr-eng \
    poppler-utils \
    libgl1-mesa-glx \
    && rm -rf /var/lib/apt/lists/*

# إعداد مجلد العمل
WORKDIR /app

# نسخ ملف المتطلبات وتثبيتها
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# نسخ باقي ملفات المشروع
COPY . .

# إنشاء مجلدات العمل إذا لم تكن موجودة
RUN mkdir -p uploads outputs

# المتغيرات البيئية (يمكن تجاوزها عند التشغيل)
ENV FLASK_APP=app__Mine_.py

# فتح المنفذ
EXPOSE 5000

# أمر التشغيل باستخدام Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app__Mine_:app"]
