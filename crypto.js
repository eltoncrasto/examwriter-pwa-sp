// crypto.js – AES-256-GCM encryption/decryption via WebCrypto
// Shared by setup.html and app.js
//
// Storage format in localStorage key "ewConfig":
//   { salt: "<base64>", iv: "<base64>", ct: "<base64>" }
//
// Key derivation: PBKDF2-SHA-256, 310,000 iterations (OWASP 2023)

const CRYPTO = (() => {
  const ENC = new TextEncoder();
  const DEC = new TextDecoder();

  function b64e(buf) {
    return btoa(String.fromCharCode(...new Uint8Array(buf)));
  }
  function b64d(str) {
    return Uint8Array.from(atob(str), c => c.charCodeAt(0));
  }

  async function deriveKey(passphrase, salt) {
    const raw = await crypto.subtle.importKey(
      "raw", ENC.encode(passphrase), "PBKDF2", false, ["deriveKey"]
    );
    return crypto.subtle.deriveKey(
      { name: "PBKDF2", salt, hash: "SHA-256", iterations: 310_000 },
      raw,
      { name: "AES-GCM", length: 256 },
      false,
      ["encrypt", "decrypt"]
    );
  }

  async function encrypt(plaintext, passphrase) {
    const salt = crypto.getRandomValues(new Uint8Array(16));
    const iv   = crypto.getRandomValues(new Uint8Array(12));
    const key  = await deriveKey(passphrase, salt);
    const ct   = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, ENC.encode(plaintext));
    return { salt: b64e(salt), iv: b64e(iv), ct: b64e(ct) };
  }

  async function decrypt(stored, passphrase) {
    const salt = b64d(stored.salt);
    const iv   = b64d(stored.iv);
    const ct   = b64d(stored.ct);
    const key  = await deriveKey(passphrase, salt);
    const pt   = await crypto.subtle.decrypt({ name: "AES-GCM", iv }, key, ct);
    return DEC.decode(pt);
  }

  return { encrypt, decrypt };
})();
