# Modernized Telegram Rating Bot

## Yangi imkoniyatlar
- 1–5 yulduzli rating tizimi
- User registration: ism, familiya, telefon contact orqali
- Bir foydalanuvchi bir o‘qituvchini bir marta baholaydi; qayta tanlasa bahosi yangilanadi
- O‘qituvchi statistikasi: average_rating, total_votes, participation_rate, final_score
- Kafedra reytingi: o‘qituvchilar final_score o‘rtachasi
- Shikoyat flow: kafedra → o‘qituvchi → matn → status
- Admin dashboard, teacher management, student_count kiritish/o‘zgartirish
- Excel export: Users, Teachers, Ratings, Complaints, Departments, Teacher Stats
- Manual va automatic daily backup
- Eski database uchun migration va backward compatibility

## ENV
- BOT_TOKEN majburiy
- ADMIN_IDS: vergul bilan ajratilgan Telegram IDlar
- CHANNEL_USERNAME default: @Qashqadaryo_PMM
- DATA_DIR default: /app/data

## Ishga tushirish
```bash
pip install -r requirements.txt
python main.py
```
