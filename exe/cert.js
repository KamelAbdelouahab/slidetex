/*
 * SlideTeX.exe — per-machine HTTPS certificate for https://localhost.
 *
 * A fresh CA is generated on each computer (never shipped), and it is restricted with X.509
 * Name Constraints to localhost / 127.0.0.1 / ::1: even if its key leaked, it could not be used
 * to impersonate any other site.
 */
"use strict";

const fs = require("fs");
const path = require("path");
const os = require("os");
const forge = require("node-forge");

const { pki, asn1, md } = forge;

function randomSerial() {
  // Positive 16-byte serial.
  const b = forge.random.getBytesSync(16);
  return "00" + forge.util.bytesToHex(b).slice(2) + "";
}

/** DER of NameConstraints { permittedSubtrees: dNSName localhost, iPAddress 127.0.0.1/32, ::1/128 } */
function nameConstraintsDer() {
  const subtree = (generalName) => asn1.create(asn1.Class.UNIVERSAL, asn1.Type.SEQUENCE, true, [generalName]);
  const dns = (name) => asn1.create(asn1.Class.CONTEXT_SPECIFIC, 2, false, name);
  const ip = (bytes) => asn1.create(asn1.Class.CONTEXT_SPECIFIC, 7, false, bytes);
  const v4 = String.fromCharCode(127, 0, 0, 1) + String.fromCharCode(255, 255, 255, 255);
  const v6 = String.fromCharCode(...new Array(15).fill(0), 1) + String.fromCharCode(...new Array(16).fill(255));
  const permitted = asn1.create(asn1.Class.CONTEXT_SPECIFIC, 0, true, [subtree(dns("localhost")), subtree(ip(v4)), subtree(ip(v6))]);
  return asn1.toDer(asn1.create(asn1.Class.UNIVERSAL, asn1.Type.SEQUENCE, true, [permitted])).getBytes();
}

function makeCa() {
  const keys = pki.rsa.generateKeyPair(2048);
  const cert = pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = randomSerial();
  cert.validity.notBefore = new Date(Date.now() - 24 * 3600e3);
  cert.validity.notAfter = new Date(Date.now() + 10 * 365 * 24 * 3600e3);
  const attrs = [
    { name: "commonName", value: `SlideTeX Local CA (${os.hostname()})` },
    { name: "organizationName", value: "SlideTeX (local only)" },
  ];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);
  cert.setExtensions([
    { name: "basicConstraints", cA: true, pathLenConstraint: 0, critical: true },
    { name: "keyUsage", keyCertSign: true, cRLSign: true, critical: true },
    { name: "subjectKeyIdentifier" },
    { id: "2.5.29.30", critical: true, value: nameConstraintsDer() },
  ]);
  cert.sign(keys.privateKey, md.sha256.create());
  return { cert, keys };
}

function makeLeaf(ca) {
  const keys = pki.rsa.generateKeyPair(2048);
  const cert = pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = randomSerial();
  cert.validity.notBefore = new Date(Date.now() - 24 * 3600e3);
  cert.validity.notAfter = new Date(Date.now() + 397 * 24 * 3600e3);
  cert.setSubject([{ name: "commonName", value: "localhost" }]);
  cert.setIssuer(ca.cert.subject.attributes);
  cert.setExtensions([
    { name: "basicConstraints", cA: false, critical: true },
    { name: "keyUsage", digitalSignature: true, keyEncipherment: true, critical: true },
    { name: "extKeyUsage", serverAuth: true },
    {
      name: "subjectAltName",
      altNames: [
        { type: 2, value: "localhost" },
        { type: 7, ip: "127.0.0.1" },
        { type: 7, ip: "::1" },
      ],
    },
    { name: "subjectKeyIdentifier" },
    { name: "authorityKeyIdentifier", keyIdentifier: ca.cert.generateSubjectKeyIdentifier().getBytes() },
  ]);
  cert.sign(ca.keys.privateKey, md.sha256.create());
  return { cert, keys };
}

/**
 * Ensure dir holds ca.crt / ca.key / localhost.crt / localhost.key (regenerating the leaf when it
 * is about to expire). Returns { caPath, caCreated, tls: {key, cert} }.
 */
function ensureCertificates(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const p = (f) => path.join(dir, f);
  let ca;
  let caCreated = false;
  try {
    ca = {
      cert: pki.certificateFromPem(fs.readFileSync(p("ca.crt"), "utf8")),
      keys: { privateKey: pki.privateKeyFromPem(fs.readFileSync(p("ca.key"), "utf8")) },
    };
    if (ca.cert.validity.notAfter.getTime() < Date.now() + 30 * 24 * 3600e3) throw new Error("CA expiring");
  } catch (_) {
    ca = makeCa();
    caCreated = true;
    fs.writeFileSync(p("ca.crt"), pki.certificateToPem(ca.cert));
    fs.writeFileSync(p("ca.key"), pki.privateKeyToPem(ca.keys.privateKey), { mode: 0o600 });
    try { fs.unlinkSync(p("localhost.crt")); } catch (_) {}
  }
  let leafOk = false;
  try {
    const leaf = pki.certificateFromPem(fs.readFileSync(p("localhost.crt"), "utf8"));
    leafOk = leaf.validity.notAfter.getTime() > Date.now() + 14 * 24 * 3600e3 && ca.cert.verify(leaf);
  } catch (_) {}
  if (!leafOk) {
    const leaf = makeLeaf(ca);
    fs.writeFileSync(p("localhost.crt"), pki.certificateToPem(leaf.cert));
    fs.writeFileSync(p("localhost.key"), pki.privateKeyToPem(leaf.keys.privateKey), { mode: 0o600 });
  }
  const caPem = fs.readFileSync(p("ca.crt"), "utf8");
  return {
    caPath: p("ca.crt"),
    caCreated,
    caThumbprint: thumbprint(pki.certificateFromPem(caPem)),
    tls: {
      key: fs.readFileSync(p("localhost.key"), "utf8"),
      cert: fs.readFileSync(p("localhost.crt"), "utf8") + caPem,
    },
  };
}

function thumbprint(cert) {
  const der = asn1.toDer(pki.certificateToAsn1(cert)).getBytes();
  return md.sha1.create().update(der).digest().toHex().toUpperCase();
}

module.exports = { ensureCertificates };
