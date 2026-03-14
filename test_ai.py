import google.generativeai as genai

# ⚠️ เอา API KEY ของจริงที่คุณพี่ได้จากเว็บ Google AI Studio มาใส่ในเครื่องหมายคำพูดนะครับ
GEMINI_API_KEY = "AIzaSyBhcoxnVIXRk-mLZUlqAuE3RZMQ5jxqEE0"
genai.configure(api_key=GEMINI_API_KEY)

print("กำลังค้นหาโมเดล AI ที่คุณสามารถใช้งานได้...")
print("-" * 40)

try:
    # สั่งให้ Google ลิสต์รายชื่อ AI ทั้งหมดที่เรามีสิทธิ์ใช้
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(f"✅ เจอโมเดลชื่อ: {m.name}")
            
    print("-" * 40)
    print("คัดลอกชื่อโมเดลด้านบน (เอาเฉพาะคำหลัง / เช่น gemini-1.5-flash) ไปใส่ใน app.py ได้เลยครับ!")
    
except Exception as e:
    print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ: {e}")
    print("💡 คำแนะนำ: ตรวจสอบให้แน่ใจว่า API Key ถูกต้อง และโปรเจกต์ใน Google AI Studio ถูกเปิดใช้งานแล้ว")