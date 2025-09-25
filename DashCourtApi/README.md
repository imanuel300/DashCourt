# DashCourtApi Backend

זהו פרויקט ה-Backend של DashCourt, שנבנה באמצעות ASP.NET Core Web API. ה-API מספק נתוני דמה מקבצי אקסל שסופקו.

## מבנה הפרויקט

- `Controllers/`: מכיל את הקונטרולרים (לדוגמה, `DataController`) שמטפלים בבקשות HTTP וחושפים את נקודות הקצה של ה-API.
- `Models/`: מכיל את המודלים של הנתונים (לדוגמה, `CRModel`, `AVGOModel`) המייצגים את מבנה הנתונים מקבצי האקסל.
- `Services/`: מכיל את השירותים (לדוגמה, `ExcelDataService`) המטפלים בלוגיקה העסקית, כגון קריאת נתונים מקבצי אקסל והמרתם למודלים.
- `Program.cs`: קובץ הכניסה ליישום, שמגדיר את שירותי ה-DI (Dependency Injection), CORS ו-Middleware.
- `DashCourtApi.csproj`: קובץ הפרויקט של .NET, המגדיר תלויות וחבילות (כגון EPPlus).

## התקנה

לפני הרצת הפרויקט, ודא שהתקנת את הכלים הבאים:

- **.NET SDK:** ודא שהתקנת את הגרסה המתאימה של .NET SDK (מומלץ .NET 8 או חדש יותר). תוכל להוריד אותו מהאתר הרשמי של Microsoft.
- **קבצי האקסל:** ודא שקבצי האקסל הבאים נמצאים בתיקיית הבסיס של הפרויקט הראשי (`D:\www\DashCourt`):
  - `CR.xlsx`
  - `AVGO.xlsx`
  - `SIT.xlsx`
  - `Inv3.xlsx`

1. **נווט לתיקיית הפרויקט של ה-API:**
   ```bash
   cd backend/DashCourtApi
   ```

2. **שחזר תלויות:**
   ```bash
   dotnet restore
   ```

   (הערה: חבילת `EPPlus` כבר נוספה בעבר, כך שייתכן ששלב זה כבר בוצע).

## הרצת הפרויקט

לאחר ההתקנה, תוכל להריץ את ה-API באמצעות הפקודה הבאה מתוך תיקיית הפרויקט (`backend/DashCourtApi`):

```bash
dotnet run
```

ה-API יופעל בדרך כלל בכתובת `https://localhost:7082` (או פורט אחר כפי שמוגדר ב-`Properties/launchSettings.json`).

## שימוש ב-API (נקודות קצה)

ה-API חושף את נקודות הקצה הבאות כדי לקבל את נתוני הדמה:

- **נתוני CR:**
  - `GET /Data/cr`
  - דוגמה: `https://localhost:7082/Data/cr`

- **נתוני AVGO:**
  - `GET /Data/avgo`
  - דוגמה: `https://localhost:7082/Data/avgo`

- **נתוני SIT:**
  - `GET /Data/sit`
  - דוגמה: `https://localhost:7082/Data/sit`

- **נתוני Inv3:**
  - `GET /Data/inv3`
  - דוגמה: `https://localhost:7082/Data/inv3`

אתה יכול לבדוק את נקודות הקצה הללו באמצעות דפדפן אינטרנט, כלי כמו Postman, או ישירות מה-Frontend שלך.
