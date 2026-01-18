# 📖 USER MANUAL
## Multi-Hostel Attendance App
**Version 6.0 | Last Updated: January 2026**

---

## 1. APP OVERVIEW

### What is this App?
This is a **Multi-Hostel Attendance Management System** designed for hostels, residential schools, and student accommodation facilities. It allows hostel wardens and administrators to track student attendance for various activities like Yoga, Mess (Day/Night), and Night Shifts.

### Who Should Use This App?
- Hostel Wardens / Rectors
- Accommodation Administrators  
- Student Coordinators
- Hostel Management Staff

### Key Features
- ✅ Multi-hostel support (each hostel has separate data)
- ✅ Four attendance categories: Yoga, Mess Day, Mess Night, Night Shift
- ✅ Real-time sync to Google Sheets
- ✅ Multi-device sync (data shared across devices)
- ✅ Admin Panel with Rankings & Reports
- ✅ Data backup and restore
- ✅ Works offline with auto-sync

---

## 2. GETTING STARTED

### First Time Registration
1. Open the app in your browser
2. Click **"Register New Hostel"**
3. Fill in the registration form:
   - Hostel Name
   - Rector/Warden Name
   - Email Address
   - Password (minimum 8 characters)
4. Click **"Register"**
5. Check your email for **OTP (One-Time Password)**
6. Enter the 6-digit OTP in the app
7. Registration complete! You will be logged in automatically

### Login (Returning Users)
1. Open the app
2. Enter your **Email** or **Hostel ID**
3. Enter your **Password**
4. Click **"Login"**
5. You will see your hostel's attendance dashboard

### Forgot Password
1. Click **"Forgot Password?"** on login screen
2. Enter your registered email
3. Click **"Send OTP"**
4. Check email and enter the OTP
5. Create a new password
6. Login with new password

---

## 3. MAIN SCREEN GUIDE

### Top Bar
- **Menu Button (☰)**: Opens the side navigation drawer
- **Search Bar**: Search for students by name
- **Avatar**: Click to view account information

### Stats Card
Shows today's attendance summary:
- Total students
- Present count
- Absent count
- Leave count

### Date Picker
- Shows current date by default
- Click to select a different date
- Note: Past synced data cannot be modified

### Student List
- Each student card shows: Name, Application Number, Status
- **Tap on student** → Cycles status: Present → Absent → Leave → Present
- Status colors:
  - 🟢 Green = Present
  - 🔴 Red = Absent
  - 🟡 Yellow = Leave

### Bottom Buttons
- **Mark All Present**: Sets all students to Present
- **Verify**: Confirms attendance is checked
- **Sync**: Uploads data to cloud (locks the data)

---

## 4. NAVIGATION DRAWER

| Menu Item | Function |
|-----------|----------|
| **Yoga** | Switch to Yoga attendance |
| **Mess Day** | Switch to Daytime mess attendance |
| **Mess Night** | Switch to Nighttime mess attendance |
| **Night Shift** | Switch to Night shift attendance |
| **Add Student** | Add a new student to the list |
| **Reset Current** | Reset current category to all Present |
| **Download CSV** | Download attendance as CSV file |
| **Quick Daily Report** | Download today's attendance report |
| **Admin Panel** | Access admin features (password required) |
| **Sync Tracker** | View sync status for recent days |
| **Logout** | Sign out from the app |

---

## 5. MARKING ATTENDANCE

### Step-by-Step Process
1. Select the correct **Date** (default is today)
2. Select the **Category** (Yoga/Mess Day/Mess Night/Night Shift)
3. Review the student list
4. **Tap each student** to change their status
5. Click **"Mark All Present"** if everyone is present
6. Click **"Verify"** to confirm
7. Click **"Sync"** to save to cloud

### Important Rules
- ⚠️ Always **Verify** before **Sync**
- ⚠️ Once synced, data is **locked** and cannot be changed
- ⚠️ Each category must be synced separately
- ✅ Data auto-saves locally as draft until synced

---

## 6. ADMIN PANEL

### Accessing Admin Panel
1. Open Navigation Drawer
2. Click **"Admin Panel"**
3. Enter Admin Password
4. Click **"Unlock"**

### Admin Panel Tabs

#### Overview Tab
- Total students count, Average attendance %, At-risk students, Top 5 performers

#### Rankings Tab
- Student leaderboard by attendance score
- Filter by category and date range
- Click any student to view details

#### AI Insights Tab
- Pattern analysis, Students at risk warnings, Recommendations

#### Reports Tab
- Section-wise Report: Select section, date range, download Excel/PDF
- Quick Daily Report: Select date, download instant report

#### Data Tab
- **Statistics**: View sheet sizes and row counts
- **Backup**: Download all data as Excel file
- **Import**: Restore data from backup file
- **Archive**: Move old data to archive sheets (safe)
- **Purge**: Permanently delete old data (careful!)

---

## 7. ERROR MESSAGES & SOLUTIONS

| Error Message | Solution |
|---------------|----------|
| "Please verify first" | Click Verify button first |
| "Already synced" | Cannot modify synced data |
| "Network error" | Check internet connection |
| "Session expired" | Login again |
| "Incorrect password" | Enter correct password |

---

## 8. DO'S AND DON'TS

### ✅ DO's
- DO verify data before syncing
- DO take regular backups
- DO keep password secure

### ❌ DON'TS
- DON'T share admin password
- DON'T close app during sync
- DON'T purge data without backup

---

## 9. FAQs

**Q: Can I edit after syncing?** A: No, synced data is locked.

**Q: Multiple devices?** A: Yes, data syncs across devices.

**Q: Is data secure?** A: Yes, each hostel has isolated Google Sheet.

**Q: Offline usage?** A: Yes, syncs when online.

---

*© 2026 Multi-Hostel Attendance App*
