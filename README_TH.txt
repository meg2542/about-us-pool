ABOUT US POOL QUEUE — V1.7.0 PUSH
=====================================

ฐานระบบ:
- Backend: V1.6.6 REALCODE ที่ใช้งานได้แล้ว
- Frontend: V1.6.7 ที่ชี้ Deployment ตัวถูกต้อง
- เพิ่ม OneSignal Web Push โดยไม่เปลี่ยน Google Sheet / Queue headers

ไฟล์ GitHub ที่ต้องอัป:
- index.html
- admin.html
- manifest.json
- OneSignalSDKWorker.js
- icon-192.png
- icon-512.png
- QR_CUSTOMER.png ใช้ตัวเดิมได้

Apps Script:
- ใช้ Code.gs ในชุดนี้แทน Code.gs เดิม
- ห้ามรัน setupPoolQueue()
- Deploy เป็น New version ของ Deployment เดิม

Script Properties ที่ต้องเพิ่มใน Apps Script:
- ONESIGNAL_APP_ID = App ID จาก OneSignal
- ONESIGNAL_REST_API_KEY = REST API Key จาก OneSignal

OneSignal Web setup:
- Web platform / Custom Code หรือ Typical Site
- Site URL / Origin: https://meg2542.github.io
- Service worker:
  path/filename: about-us-pool/OneSignalSDKWorker.js
  scope: /about-us-pool/

iPhone:
- ต้อง iOS/iPadOS 16.4+
- ต้อง Add to Home Screen
- เปิดเว็บจากไอคอนที่ Home Screen
- กด “เปิดแจ้งเตือน”
- จากนั้นกด “ทดสอบ” เพื่อตรวจ Push

Push flow:
Admin CALL_NEXT
→ เปลี่ยนคิวเป็น CALLED
→ Backend ส่ง OneSignal Push ไป external_id = poolq_<ClientToken>
→ ถ้า OneSignal ผิด/ล่ม ระบบคิวยังทำงานต่อ (Push เป็น best-effort)

ข้อสำคัญ:
- REST API Key อยู่ใน Apps Script Script Properties เท่านั้น
- ห้ามใส่ REST API Key ใน index.html / GitHub
