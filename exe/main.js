/*
 * SlideTeX.exe — self-contained installer + local server (no Node.js / npm needed).
 *
 *   SlideTeX.exe              install (or repair) for the current Windows user, then start
 *   SlideTeX.exe --serve      run the server only (used by the hidden autostart at logon)
 *   SlideTeX.exe --uninstall  remove everything
 *   SlideTeX.exe --status     print what is installed
 *
 * Everything lives in %LOCALAPPDATA%\SlideTeX — no administrator rights required.
 */
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const https = require("https");
const { execFileSync, spawn } = require("child_process");

const { createApp, listenLoopback, toolInfo, warmUp, VERSION } = require("../lib/app");
const { ensureCertificates } = require("./cert");
const ASSETS = require("../build/assets.generated.js");

const IS_WIN = process.platform === "win32";
const args = process.argv.slice(2);
const argVal = (n, d) => {
  const i = args.indexOf(n);
  return i >= 0 && args[i + 1] ? args[i + 1] : d;
};
const PORT = parseInt(argVal("--port", "41730"), 10);
const APPDIR = argVal("--data-dir", path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".local", "share"), "SlideTeX"));
const EXE_NAME = "SlideTeX.exe";
const INSTALLED_EXE = path.join(APPDIR, EXE_NAME);
const MANIFEST_PATH = path.join(APPDIR, "manifest.xml");
const VBS_PATH = path.join(APPDIR, "start-hidden.vbs");
const LOG_PATH = path.join(APPDIR, "slidetex.log");
const ADDIN_ID = "6f3c2a1e-9b4d-4c7e-8a51-2d7e0b9f4c13";
const KEY_WEF = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";
const KEY_RUN = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const KEY_UNINSTALL = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SlideTeX";

// ------------------------------------------------------------------ embedded assets
const bufCache = new Map();
function asset(p) {
  const a = ASSETS.files[p];
  if (!a) return null;
  if (!bufCache.has(p)) bufCache.set(p, { type: a[0], buf: Buffer.from(a[1], "base64") });
  return bufCache.get(p);
}
function mountStatic(app) {
  app.use((req, res, next) => {
    if (req.method !== "GET" && req.method !== "HEAD") return next();
    let p = decodeURIComponent(req.path);
    if (p === "/") p = "/taskpane.html";
    const a = asset(p);
    if (!a) return next();
    res.setHeader("Content-Type", a.type);
    res.end(req.method === "HEAD" ? undefined : a.buf);
  });
}

let resvgReady = null;
async function renderPng(svg, px) {
  const resvg = require("@resvg/resvg-wasm");
  if (!resvgReady) resvgReady = resvg.initWasm(asset("/__resvg.wasm").buf);
  await resvgReady;
  const r = new resvg.Resvg(svg, { fitTo: { mode: "width", value: px }, background: "rgba(0,0,0,0)" });
  return r.render().asPng();
}

// ------------------------------------------------------------------ helpers
const say = (s = "") => process.stdout.write(s + "\n");
function sh(cmd, argv, opts = {}) {
  try {
    const out = execFileSync(cmd, argv, { stdio: ["ignore", "pipe", "pipe"], windowsHide: true, ...opts });
    return { ok: true, out: String(out || "") };
  } catch (e) {
    return { ok: false, out: String((e.stdout || "") + (e.stderr || "")), code: e.status };
  }
}
const reg = (...a) => sh("reg.exe", a);

/** Make MiKTeX / TeX Live reachable even if PATH was not refreshed (login-time autostart). */
function extendPathForTex() {
  if (!IS_WIN) return;
  const la = process.env.LOCALAPPDATA || "";
  const pf = process.env.ProgramFiles || "C:\\Program Files";
  const cands = [
    path.join(la, "Programs", "MiKTeX", "miktex", "bin", "x64"),
    path.join(pf, "MiKTeX", "miktex", "bin", "x64"),
    path.join(pf, "MiKTeX 2.9", "miktex", "bin", "x64"),
  ];
  try {
    for (const y of fs.readdirSync("C:\\texlive").filter((d) => /^\d{4}$/.test(d)).sort().reverse()) {
      cands.push(`C:\\texlive\\${y}\\bin\\windows`, `C:\\texlive\\${y}\\bin\\win64`, `C:\\texlive\\${y}\\bin\\win32`);
    }
  } catch (_) {}
  const key = Object.keys(process.env).find((k) => k.toLowerCase() === "path") || "PATH";
  const cur = (process.env[key] || "").split(";");
  const add = cands.filter((c) => fs.existsSync(c) && !cur.some((x) => x.toLowerCase() === c.toLowerCase()));
  if (add.length) process.env[key] = cur.concat(add).join(";");
}

function health(timeoutMs = 1500) {
  return new Promise((resolve) => {
    const req = https.get(
      { host: "127.0.0.1", port: PORT, path: "/api/health", rejectUnauthorized: false, timeout: timeoutMs },
      (res) => {
        let d = "";
        res.on("data", (c) => (d += c));
        res.on("end", () => {
          try {
            resolve(JSON.parse(d));
          } catch (_) {
            resolve(null);
          }
        });
      }
    );
    req.on("error", () => resolve(null));
    req.on("timeout", () => {
      req.destroy();
      resolve(null);
    });
  });
}

function killOtherInstances() {
  if (!IS_WIN) return;
  sh("taskkill.exe", ["/F", "/IM", EXE_NAME, "/FI", `PID ne ${process.pid}`]);
}

// ------------------------------------------------------------------ server
async function serve() {
  fs.mkdirSync(APPDIR, { recursive: true });
  if (!process.stdout.isTTY) {
    const log = fs.createWriteStream(LOG_PATH, { flags: "a" });
    const w = (...a) => log.write(new Date().toISOString() + " " + a.join(" ") + "\n");
    console.log = w;
    console.error = w;
  }
  const h = await health();
  if (h && h.name === "SlideTeX") {
    console.log("SlideTeX already running on port " + PORT);
    return false;
  }
  extendPathForTex();
  const certs = ensureCertificates(path.join(APPDIR, "certs"));
  const app = createApp({ port: PORT, mountStatic, renderPng });
  await listenLoopback(app, PORT, certs.tls);
  const info = await toolInfo();
  console.log(`SlideTeX ${VERSION} listening on https://localhost:${PORT} — LaTeX: ${info.latex ? info.latexVersion : "absent"}`);
  warmUp((m) => console.log(m)).catch(() => {});
  return true;
}

// ------------------------------------------------------------------ install
function writeManifest() {
  const xml = ASSETS.manifestTemplate.split("{{BASE}}").join(`https://localhost:${PORT}`).split("{{ORIGIN}}").join(`https://localhost:${PORT}`);
  fs.writeFileSync(MANIFEST_PATH, xml);
}

function caInstalled(thumb) {
  return sh("certutil.exe", ["-user", "-store", "Root", thumb]).ok;
}

async function install() {
  if (!IS_WIN) {
    say("L'installation automatique est réservée à Windows. Utilisez --serve sur les autres systèmes.");
    return 1;
  }
  sh("cmd.exe", ["/c", "chcp 65001 >nul"]);
  say(`SlideTeX ${VERSION} — installation pour l'utilisateur ${os.userInfo().username}`);
  say("");

  say("1/6  Arrêt d'une éventuelle ancienne version…");
  killOtherInstances();

  say("2/6  Copie dans " + APPDIR);
  fs.mkdirSync(APPDIR, { recursive: true });
  if (path.resolve(process.execPath).toLowerCase() !== path.resolve(INSTALLED_EXE).toLowerCase()) {
    let copied = false;
    for (let i = 0; i < 10 && !copied; i++) {
      try {
        fs.copyFileSync(process.execPath, INSTALLED_EXE);
        copied = true;
      } catch (e) {
        await new Promise((r) => setTimeout(r, 500));
      }
    }
    if (!copied) throw new Error("Impossible de copier l'exécutable (fichier verrouillé ?)");
  }

  say("3/6  Certificat HTTPS local (réservé à localhost)…");
  const certs = ensureCertificates(path.join(APPDIR, "certs"));
  if (!caInstalled(certs.caThumbprint)) {
    if (certs.caCreated) sh("certutil.exe", ["-user", "-delstore", "Root", `SlideTeX Local CA (${os.hostname()})`]);
    say("     ➜ Windows va afficher un avertissement de sécurité : cliquez sur « Oui ».");
    const r = sh("certutil.exe", ["-user", "-addstore", "Root", certs.caPath]);
    if (!r.ok || !caInstalled(certs.caThumbprint)) {
      say("     ✗ Certificat refusé ou bloqué par une stratégie d'entreprise.");
      say("       Relancez SlideTeX.exe et acceptez l'avertissement, ou demandez à votre service informatique.");
      return 1;
    }
  }

  say("4/6  Enregistrement du complément dans PowerPoint…");
  writeManifest();
  const r1 = reg("add", KEY_WEF, "/v", ADDIN_ID, "/t", "REG_SZ", "/d", MANIFEST_PATH, "/f");
  if (!r1.ok) throw new Error("Échec de l'écriture du registre : " + r1.out);

  say("5/6  Démarrage automatique à l'ouverture de session…");
  fs.writeFileSync(
    VBS_PATH,
    `' SlideTeX - starts the local server without a console window\r\n` +
      `CreateObject("WScript.Shell").Run """${INSTALLED_EXE}"" --serve", 0, False\r\n`
  );
  reg("add", KEY_RUN, "/v", "SlideTeX", "/t", "REG_SZ", "/d", `wscript.exe "${VBS_PATH}"`, "/f");
  reg("add", KEY_UNINSTALL, "/v", "DisplayName", "/t", "REG_SZ", "/d", "SlideTeX - LaTeX pour PowerPoint", "/f");
  reg("add", KEY_UNINSTALL, "/v", "DisplayVersion", "/t", "REG_SZ", "/d", VERSION, "/f");
  reg("add", KEY_UNINSTALL, "/v", "Publisher", "/t", "REG_SZ", "/d", "SlideTeX", "/f");
  reg("add", KEY_UNINSTALL, "/v", "DisplayIcon", "/t", "REG_SZ", "/d", INSTALLED_EXE, "/f");
  reg("add", KEY_UNINSTALL, "/v", "InstallLocation", "/t", "REG_SZ", "/d", APPDIR, "/f");
  reg("add", KEY_UNINSTALL, "/v", "UninstallString", "/t", "REG_SZ", "/d", `"${INSTALLED_EXE}" --uninstall`, "/f");
  reg("add", KEY_UNINSTALL, "/v", "NoModify", "/t", "REG_DWORD", "/d", "1", "/f");
  reg("add", KEY_UNINSTALL, "/v", "NoRepair", "/t", "REG_DWORD", "/d", "1", "/f");

  say("6/6  Lancement du serveur local…");
  spawn("wscript.exe", [VBS_PATH], { detached: true, stdio: "ignore", windowsHide: true }).unref();
  let h = null;
  for (let i = 0; i < 40 && !(h && h.name === "SlideTeX"); i++) {
    await new Promise((r) => setTimeout(r, 500));
    h = await health();
  }
  if (!h) {
    say(`     ✗ Le serveur ne répond pas (port ${PORT} occupé ?). Journal : ${LOG_PATH}`);
    return 1;
  }

  say("");
  say("✔ SlideTeX est installé.");
  say("");
  say("  ➜ Fermez puis rouvrez PowerPoint.");
  say("  ➜ Onglet Accueil › groupe « LaTeX » › Équation LaTeX.");
  say("    (Si le bouton n'apparaît pas : Insertion › Mes compléments › Compléments de développement › SlideTeX.)");
  say("");
  say(`  Moteur LaTeX local : ${h.latex && h.dvisvgm ? "détecté (" + h.latexVersion + ")" : "non détecté — le moteur MathJax intégré suffit pour les équations ; installez MiKTeX pour TikZ & co."}`);
  say("  Désinstaller : Paramètres Windows › Applications › SlideTeX, ou « SlideTeX.exe --uninstall ».");
  return 0;
}

async function uninstall() {
  if (!IS_WIN) return 0;
  sh("cmd.exe", ["/c", "chcp 65001 >nul"]);
  say("Désinstallation de SlideTeX…");
  killOtherInstances();
  reg("delete", KEY_WEF, "/v", ADDIN_ID, "/f");
  reg("delete", KEY_RUN, "/v", "SlideTeX", "/f");
  reg("delete", KEY_UNINSTALL, "/f");
  try {
    const { caThumbprint } = ensureCertificatesReadOnly();
    if (caThumbprint) sh("certutil.exe", ["-user", "-delstore", "Root", caThumbprint]);
  } catch (_) {}
  // The running exe may live in APPDIR: delete the folder once we have exited.
  // Retries for up to a minute because this very exe may still be open (window waiting for Enter).
  const loop = `for /l %i in (1,1,60) do @(ping 127.0.0.1 -n 2 >nul & rmdir /s /q "${APPDIR}" 2>nul & if not exist "${APPDIR}" exit)`;
  spawn("cmd.exe", ["/d", "/c", loop], { detached: true, stdio: "ignore", windowsHide: true, windowsVerbatimArguments: true }).unref();
  say("✔ SlideTeX a été supprimé. Redémarrez PowerPoint.");
  return 0;
}

function ensureCertificatesReadOnly() {
  const forge = require("node-forge");
  const pem = fs.readFileSync(path.join(APPDIR, "certs", "ca.crt"), "utf8");
  const cert = forge.pki.certificateFromPem(pem);
  const der = forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes();
  return { caThumbprint: forge.md.sha1.create().update(der).digest().toHex().toUpperCase() };
}

async function status() {
  const h = await health();
  say(`SlideTeX ${VERSION}`);
  say(`  Dossier   : ${APPDIR} ${fs.existsSync(APPDIR) ? "" : "(absent)"}`);
  say(`  Serveur   : ${h ? `actif sur https://localhost:${PORT} (v${h.version || "?"})` : "arrêté"}`);
  if (IS_WIN) say(`  PowerPoint: ${reg("query", KEY_WEF, "/v", ADDIN_ID).ok ? "complément enregistré" : "non enregistré"}`);
  if (h) say(`  LaTeX     : ${h.latex ? h.latexVersion : "non détecté"}`);
  return 0;
}

function pause() {
  if (!process.stdin.isTTY) return Promise.resolve();
  say("");
  process.stdout.write("Appuyez sur Entrée pour fermer cette fenêtre…");
  return new Promise((r) => process.stdin.once("data", r));
}

// ------------------------------------------------------------------ main
(async () => {
  let code = 0;
  try {
    if (args.includes("--serve")) {
      const started = await serve();
      if (!started) process.exit(0);
      return; // keep running
    } else if (args.includes("--uninstall")) code = await uninstall();
    else if (args.includes("--status")) code = await status();
    else code = await install();
  } catch (e) {
    say("✗ Erreur : " + (e && e.message ? e.message : e));
    code = 1;
  }
  if (!args.includes("--quiet")) await pause();
  process.exit(code);
})();
