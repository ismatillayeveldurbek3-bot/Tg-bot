# MENU / BACK FIX

Ushbu versiyada tugmalar iyerarxiyasi va back tugmalari mantiqan tartiblandi.

## Asosiy o‘zgarishlar
- `admin:teachers` callback handleri tiklandi. Oldingi holatda Admin paneldagi “O‘qituvchilar” tugmasi fallbackga tushishi mumkin edi.
- Sozlamalar ichidagi ichki oynalarda `⬅️ Orqaga` tugmalari to‘g‘ri ota bo‘limga qaytaradigan qilindi.
- Sozlamalar menyusidan takroriy `Shikoyatlar` tugmasi olib tashlandi, chunki admin panelda alohida `Shikoyatlar` bo‘limi bor.
- Kafedra sozlamalari ichidagi ikki xil back tugma bitta mantiqiy `⬅️ Orqaga` tugmasiga qisqartirildi.
- Tozalash tasdiqlash oynasida `Bekor` va `Tozalash menyusi` takrori bitta `⬅️ Orqaga`ga qisqartirildi.
- O‘qituvchi qo‘shish/tahrirlash/o‘quvchilar sonini kiritishdan keyin admin bosh menyusiga sakramaydi, tegishli o‘qituvchi yoki kafedra oynasiga qaytadi.
- Admin qo‘shish/olib tashlashdan keyin root admin panelga emas, Adminlar sozlamasiga qaytadi.
- Ishlatilmay qolgan `url:noop` callback uchun xavfsiz handler qo‘shildi.

## Tekshiruv
- `python -m py_compile main.py` muvaffaqiyatli o‘tdi.
