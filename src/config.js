// src/config.js
export const SELECTORS = {
  userName: '#UserName',
  password: '#Password',
  language: '#Language',
  loginDate: '#LoginDataTime',
  validateCaptcha: '#ValidateCaptcha',
  captchaImage: '#imgcapt',
  loginButton: 'button[type="submit"], button[onclick*="ValidateAndLogin"], #btnView, .login-btn',
  // follow-up form example (change to match actual page)
  nextFormSelector: 'form#profile-form',
  nextField1: 'input[name="field1"]',
  nextField2: 'input[name="field2"]',
  nextSubmit: 'button[type="submit"]'
}
