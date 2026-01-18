# 🔗 FOOTER & NAVIGATION GUIDE
## Multi-Hostel Attendance App - Web Integration

---

## CLICKABLE POLICY PAGES CREATED

| Page | File | URL Path |
|------|------|----------|
| Privacy Policy | `privacy.html` | `/privacy.html` |
| Terms & Conditions | `terms.html` | `/terms.html` |
| Disclaimer | `disclaimer.html` | `/disclaimer.html` |
| Contact Us | `contact.html` | `/contact.html` |
| Help/User Guide | `help.html` | `/help.html` |

---

## FOOTER HTML CODE

Add this to your main `index.html` or `welcome.html`:

```html
<footer style="
    background: #f5f5f7;
    padding: 20px;
    text-align: center;
    font-size: 12px;
    color: #86868b;
    margin-top: 20px;
    border-top: 1px solid #e5e5e7;
">
    <div style="margin-bottom: 12px;">
        <a href="privacy.html" style="color: #004B8E; margin: 0 10px; text-decoration: none;">Privacy Policy</a>
        <a href="terms.html" style="color: #004B8E; margin: 0 10px; text-decoration: none;">Terms & Conditions</a>
        <a href="disclaimer.html" style="color: #004B8E; margin: 0 10px; text-decoration: none;">Disclaimer</a>
        <a href="contact.html" style="color: #004B8E; margin: 0 10px; text-decoration: none;">Contact Us</a>
        <a href="help.html" style="color: #004B8E; margin: 0 10px; text-decoration: none;">Help</a>
    </div>
    <div>
        © 2026 Multi-Hostel Attendance App. All rights reserved.
    </div>
</footer>
```

---

## NAVIGATION DRAWER LINKS

Add to side menu/drawer:

```html
<div class="drawer-section" style="border-top: 1px solid #e5e5e7; margin-top: 16px; padding-top: 16px;">
    <div class="drawer-item" onclick="window.location.href='help.html'">
        <span class="material-symbols-outlined">help</span>
        <span>Help & Guide</span>
    </div>
    <div class="drawer-item" onclick="window.location.href='contact.html'">
        <span class="material-symbols-outlined">support_agent</span>
        <span>Contact Support</span>
    </div>
    <div class="drawer-item" onclick="window.location.href='privacy.html'">
        <span class="material-symbols-outlined">privacy_tip</span>
        <span>Privacy Policy</span>
    </div>
</div>
```

---

## WELCOME PAGE LINKS

Add to login/welcome page footer:

```html
<div style="margin-top: 24px; text-align: center; font-size: 12px; color: #86868b;">
    By using this app, you agree to our 
    <a href="terms.html" style="color: #004B8E;">Terms</a> and 
    <a href="privacy.html" style="color: #004B8E;">Privacy Policy</a>
</div>
```

---

## ANCHOR TEXT SUGGESTIONS

| Link | Display Text |
|------|--------------|
| Privacy Policy | "Privacy Policy" or "🔒 Privacy" |
| Terms & Conditions | "Terms & Conditions" or "Terms of Use" |
| Disclaimer | "Disclaimer" or "Legal Notice" |
| Contact Us | "Contact Us" or "📞 Support" |
| Help | "Help" or "📖 User Guide" |

---

## MOBILE-FIRST FOOTER (Compact)

For mobile view, use icons:

```html
<div style="display: flex; justify-content: center; gap: 16px; padding: 16px;">
    <a href="privacy.html" title="Privacy">🔒</a>
    <a href="terms.html" title="Terms">📜</a>
    <a href="disclaimer.html" title="Disclaimer">⚠️</a>
    <a href="contact.html" title="Contact">📞</a>
    <a href="help.html" title="Help">❓</a>
</div>
```

---

## FILES CHECKLIST

✅ `privacy.html` - Privacy Policy
✅ `terms.html` - Terms & Conditions  
✅ `disclaimer.html` - Disclaimer
✅ `contact.html` - Contact Us (WhatsApp: +91 8390229604)
✅ `help.html` - User Manual/Help

All pages have:
- Consistent styling (Inter font, VSS brand colors)
- Back to App link
- Footer navigation to other pages
- Mobile-responsive design
- India-compliant content

---

## NETLIFY DEPLOYMENT

All HTML files in root folder will automatically work with Netlify:
- `yoursite.netlify.app/privacy.html`
- `yoursite.netlify.app/terms.html`
- etc.

No additional configuration needed!
