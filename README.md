## EGAT SBK Worklist Convert
โปรแกรม Python สำหรับแปลง Worklist (Excel File) โรงไฟฟ้าพระนครใต้ (EGAT) ให้สำหรับให้กับ Template Maximo (Excel File)
การทำงานของโปรแกรม 
- จะเป็นการย้าย Cell
- Split Resp. Cell (กรณี 1 Row มี 2 หน่วยงาน)
- แปลง Respone Carft
- เพิ่ม TYPE ของหน่วยงาน
- Split Acitivity
- Highlight cell ที่เกิน 85 ตัวอักษร (ตามข้ำจำกัดของโปรแกรม Maximo)
- จัดลำดับ TASK order

## ขั้นตอนก่อนใช้งาน

1. ทำการติดตั้ง Python Version >= 3.9 https://www.python.org/downloads/

![image](https://github.com/user-attachments/assets/1f4ea6fd-9728-4a36-8a73-09e5e978a6bc)

2. ให้ติดตั้ง libary Python โดยเปิด Command Promt แล้วใช้คำสั่งดังนี้

```
pip install openpyxl
```
![image](https://github.com/user-attachments/assets/c218fdc6-8ede-4434-9ccb-6622f54ed635)
  
3. Download File ของโปรแกรมนี้ โดยกดปุ่มสีเขียว (Code) https://github.com/PiyaRawing/EGAT_SBK_Worklist_Convert

4. เลือก Download ZIP
 
5. ทำการแตกไฟล์ .zip จะได้ Folder ที่ชื่อว่า EGAT_SBK_Worklist_Convert-main มา

![image](https://github.com/user-attachments/assets/11a4eed2-36da-499b-bfcc-c775408c0d9c)

## ขั้นตอนใช้งาน
1. เปิด Folder ที่ทำการแตกไฟล์ไว้ 
 
2. Double Click ที่ไฟล์ **main.py**

3. ทำการกดเลือกไฟล์ Worklist.xlsx

4. ติ๊ก CheckBox หากต้องการให้โปรแกรม Highlight สี ที่ Cell ที่มีตัวอักษรเกิน 100 ตัว

5. เลือก Sheet ที่ต้องการจะ Convert

6. กดปุ่ม Convert to Maximo

7. เลือกที่อยู่ที่จะ Save File

8. เสร็จสิ้นการทำงาน

## ปัญหาที่อาจพบเจอ
1.เมื่อกด เลือกไฟล์ หรือ ปุ่ม Convert แล้วพบ Error แจ้งเตือน เกิดจาก **เปิดไฟล์ที่จะทำการ** Convert ไว้อยู่ ให้ปิดไฟล์นั้นก่อน

2.เมื่อกด Save แล้วพบ Error แจ้งเตือน เกิดจาก ไฟล์ชื่อเดียวกัน เปิดไว้อยู่ **ให้ปิดไฟล์นั้นก่อน** เพราะ โปรแกรมจะ Save ไฟล์ทับขณะไฟล์นั้นถูกเปิดไว้อยู่ไม่ได้

![image](https://github.com/user-attachments/assets/2e66d0bb-17a0-416f-9a6c-1a2fdc32235c)
 
> [!CAUTION]
> **File main.py** และ **Respone - Do not Delete.xlsx** ต้องอยู่ภายใต้ Folder เดียวกัน

> [!CAUTION]
> การแยก Activity จะแยกด้วย **เลขข้อจุด** (1.aaa2.bbb) หรือหากมีเครื่องหมาย วงเล็บ () **ต้องใส่ให้ครบทุกครั้ง** เพราะโปรแกรมจะมองสิ่งที่อยู่ในวงเว็บเป็นเนื้อหาของข้อนั้น เครื่องหมายอื่นใส่ได้ปกติไม่มีผลต่อการแยก
 
![image](https://github.com/user-attachments/assets/b66c8e6d-de17-4987-b149-4b3580b6a4e5)
