// ============================================================
//  GOLD SERVICE DETAILING — AUTO BOOKING CONFIRMATION EMAIL
//  Paste this into your "GSD — Booking Requests" Apps Script
//  Extensions → Apps Script → replace everything → Save → Deploy
// ============================================================

// ---- CONFIGURATION — edit these values ----
const CONFIG = {
  businessName:  "Gold Service Detailing",
  ownerName:     "Gold Service Detailing Team",
  ownerEmail:    "Goldservicedetailing@gmail.com",
  phone:         "(732) 430-6468",
  website:       "www.goldservicedetailing.com",
  instagram:     "@goldservicedetailing",
  reviewLink:    "https://g.page/r/YOUR_GOOGLE_REVIEW_LINK/review", // <-- replace with your Google review link
  bookingSheet:  "GSD — Booking Requests", // name of your sheet tab
};
// -------------------------------------------


// ============================================================
//  PART 1 — doPost()
//  Receives booking data from your website form and saves it,
//  then fires the confirmation email automatically.
// ============================================================
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp,
      data.firstName,
      data.lastName,
      data.phone,
      data.email,
      data.vehicle,
      data.package,
      data.addon,
      data.dateLocation,
      data.notes
    ]);

    // Fire confirmation email if we have an email address
    if (data.email) {
      sendConfirmationEmail(data);
    }

    // Notify YOU (the owner) about the new booking
    sendOwnerNotification(data);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
//  PART 2 — sendConfirmationEmail()
//  Sends a branded HTML confirmation to the customer.
// ============================================================
function sendConfirmationEmail(data) {
  var firstName = data.firstName || "there";
  var subject   = "✅ Booking Request Received — Gold Service Detailing";

  var html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { margin: 0; padding: 0; background: #f4f4f4; font-family: Arial, sans-serif; }
    .wrapper { max-width: 600px; margin: 0 auto; background: #ffffff; }
    .header {
      background: #0D0D0D;
      padding: 32px 40px;
      text-align: center;
      border-bottom: 3px solid #C9A84C;
    }
    .header h1 {
      color: #C9A84C;
      font-size: 22px;
      font-weight: 700;
      margin: 0;
      letter-spacing: 1px;
    }
    .header p { color: #888; font-size: 12px; margin: 6px 0 0; letter-spacing: 2px; text-transform: uppercase; }
    .body { padding: 40px; }
    .greeting { font-size: 22px; font-weight: 700; color: #1A1A1A; margin-bottom: 12px; }
    .intro { font-size: 15px; color: #444; line-height: 1.6; margin-bottom: 28px; }
    .summary-box {
      background: #FAFAFA;
      border: 1px solid #E8C97A;
      border-radius: 6px;
      padding: 24px;
      margin-bottom: 28px;
    }
    .summary-box h3 {
      color: #8B6914;
      font-size: 11px;
      letter-spacing: 2px;
      text-transform: uppercase;
      margin: 0 0 16px;
    }
    .summary-row {
      display: flex;
      justify-content: space-between;
      padding: 8px 0;
      border-bottom: 1px solid #eee;
      font-size: 14px;
    }
    .summary-row:last-child { border-bottom: none; }
    .summary-label { color: #888; }
    .summary-value { color: #1A1A1A; font-weight: 600; }
    .promo-box {
      background: #FFF8E7;
      border-left: 4px solid #C9A84C;
      padding: 16px 20px;
      border-radius: 0 6px 6px 0;
      margin-bottom: 28px;
      font-size: 14px;
      color: #8B6914;
    }
    .next-steps h3 { font-size: 16px; font-weight: 700; color: #1A1A1A; margin-bottom: 12px; }
    .step { display: flex; gap: 12px; margin-bottom: 12px; align-items: flex-start; }
    .step-num {
      background: #C9A84C; color: #fff;
      width: 24px; height: 24px; border-radius: 50%;
      font-size: 12px; font-weight: 700;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0; margin-top: 1px;
    }
    .step-text { font-size: 14px; color: #444; line-height: 1.5; }
    .cta-btn {
      display: block;
      background: #C9A84C;
      color: #0D0D0D !important;
      text-decoration: none;
      text-align: center;
      padding: 14px 32px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 14px;
      letter-spacing: 1px;
      text-transform: uppercase;
      margin: 28px 0;
    }
    .footer {
      background: #0D0D0D;
      padding: 24px 40px;
      text-align: center;
    }
    .footer p { color: #666; font-size: 12px; margin: 4px 0; }
    .footer a { color: #C9A84C; text-decoration: none; }
  </style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <h1>✦ Gold Service Detailing</h1>
    <p>Premium Mobile Auto Detailing · Carteret, NJ</p>
  </div>

  <div class="body">
    <p class="greeting">Hey ${firstName}! 👋</p>
    <p class="intro">
      We've received your booking request and we're excited to make your car shine.
      We'll confirm your appointment details within <strong>24 hours</strong> — usually much faster!
    </p>

    <div class="summary-box">
      <h3>Your Booking Summary</h3>
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#888;font-size:14px;">Service</td><td style="padding:8px 0;border-bottom:1px solid #eee;color:#1A1A1A;font-weight:600;font-size:14px;text-align:right;">${data.package || "—"}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#888;font-size:14px;">Add-On</td><td style="padding:8px 0;border-bottom:1px solid #eee;color:#1A1A1A;font-weight:600;font-size:14px;text-align:right;">${data.addon || "None"}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#888;font-size:14px;">Vehicle</td><td style="padding:8px 0;border-bottom:1px solid #eee;color:#1A1A1A;font-weight:600;font-size:14px;text-align:right;">${data.vehicle || "—"}</td></tr>
        <tr><td style="padding:8px 0;color:#888;font-size:14px;">Date / Location</td><td style="padding:8px 0;color:#1A1A1A;font-weight:600;font-size:14px;text-align:right;">${data.dateLocation || "—"}</td></tr>
      </table>
    </div>

    ${data.notes && data.notes.toLowerCase().includes("first") ? `
    <div class="promo-box">
      🎉 <strong>First-time customer discount noted!</strong> Your $20 off will be applied at the time of service.
    </div>` : ""}

    <div class="next-steps">
      <h3>What Happens Next</h3>
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-text"><strong>We'll call or text you</strong> within 24 hours to confirm your exact time and address.</div>
      </div>
      <div class="step">
        <div class="step-num">2</div>
        <div class="step-text"><strong>We show up fully equipped</strong> — no water or electricity needed on your end.</div>
      </div>
      <div class="step">
        <div class="step-num">3</div>
        <div class="step-text"><strong>You get before & after photos</strong> of every job so you can see the transformation.</div>
      </div>
    </div>

    <p style="font-size:14px;color:#444;margin-top:24px;">
      Questions? Reply to this email or reach us directly:<br>
      📞 <strong>${CONFIG.phone}</strong> &nbsp;|&nbsp; 📱 <a href="https://instagram.com/goldservicedetailing" style="color:#C9A84C;">${CONFIG.instagram}</a>
    </p>

    <a href="https://${CONFIG.website}" class="cta-btn">Visit Our Website</a>

    <p style="font-size:13px;color:#888;text-align:center;">
      We can't wait to detail your vehicle.<br>
      <em>— The Gold Service Detailing Team</em>
    </p>
  </div>

  <div class="footer">
    <p><a href="https://${CONFIG.website}">${CONFIG.website}</a> &nbsp;·&nbsp; ${CONFIG.phone}</p>
    <p>100 Roosevelt Ave, Carteret, NJ 07008</p>
    <p style="margin-top:8px;color:#444;">Follow us: <a href="https://instagram.com/goldservicedetailing">Instagram</a> &nbsp;·&nbsp; <a href="https://facebook.com/goldservicedetailing">Facebook</a> &nbsp;·&nbsp; <a href="https://tiktok.com/@goldservicedetailing">TikTok</a></p>
  </div>

</div>
</body>
</html>`;

  GmailApp.sendEmail(
    data.email,
    subject,
    // Plain text fallback
    `Hi ${firstName}, we received your booking request for ${data.package || "a detail"} on ${data.dateLocation || "your requested date"}. We'll confirm within 24 hours. Questions? Call us at ${CONFIG.phone}. — Gold Service Detailing`,
    {
      name:    CONFIG.businessName,
      replyTo: CONFIG.ownerEmail,
      htmlBody: html
    }
  );
}


// ============================================================
//  PART 3 — sendOwnerNotification()
//  Sends YOU a quick SMS-style email alert for every new booking.
//  Goes to your Gmail so you never miss a lead.
// ============================================================
function sendOwnerNotification(data) {
  var subject = `🔔 New Booking Request — ${data.firstName || ""} ${data.lastName || ""} (${data.package || "Unknown Service"})`;

  var body = `
NEW BOOKING REQUEST
-------------------
Name:         ${data.firstName || ""} ${data.lastName || ""}
Phone:        ${data.phone || "—"}
Email:        ${data.email || "—"}
Vehicle:      ${data.vehicle || "—"}
Service:      ${data.package || "—"}
Add-On:       ${data.addon || "None"}
Date/Location:${data.dateLocation || "—"}
Notes:        ${data.notes || "—"}
Submitted:    ${data.timestamp || new Date().toLocaleString()}
-------------------
Reply or call to confirm within 24 hours.
  `;

  GmailApp.sendEmail(
    CONFIG.ownerEmail,
    subject,
    body,
    { name: "GSD Booking System" }
  );
}


// ============================================================
//  PART 4 — onFormSubmit() TRIGGER (for Google Forms users)
//  If you ever use Google Forms instead of your website form,
//  this trigger fires the same confirmation email.
//  Set up: Triggers → onFormSubmit → From spreadsheet → On form submit
// ============================================================
function onFormSubmit(e) {
  if (!e || !e.values) return;
  var row = e.values;
  var data = {
    timestamp:   row[0],
    firstName:   row[1],
    lastName:    row[2],
    phone:       row[3],
    email:       row[4],
    vehicle:     row[5],
    package:     row[6],
    addon:       row[7],
    dateLocation:row[8],
    notes:       row[9]
  };
  if (data.email) sendConfirmationEmail(data);
  sendOwnerNotification(data);
}
