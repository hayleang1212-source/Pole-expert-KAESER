import { useState, useEffect, useMemo, useRef } from "react";
import { createPortal } from "react-dom";
import { Document, Page, Outline, pdfjs } from "react-pdf";
import "react-pdf/dist/Page/AnnotationLayer.css";
import "react-pdf/dist/Page/TextLayer.css";
import { auth, db } from "./firebase";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signOut,
  sendPasswordResetEmail,
} from "firebase/auth";
import { doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, arrayUnion, onSnapshot, serverTimestamp, addDoc, collection, query, where } from "firebase/firestore";

const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbypM6B9C8jnhsGWDvk9u4bJi3np-Q0qYEbwvHMUQ3KPooSgF1UHoKQu1_qM5-XdsDxQeA/exec";

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1] || "");
    reader.onerror = () => reject(new Error("Lecture du fichier impossible."));
    reader.readAsDataURL(blob);
  });
}

// Taille de chaque morceau envoyé à l'Apps Script (avant encodage base64).
// 8 Mo -> ~10,7 Mo une fois encodé, largement sous la limite d'environ 50 Mo
// que supporte une requête vers un Web App Apps Script. Doit être un multiple
// de 256 Ko (contrainte de l'API Drive "resumable"), ce qui est le cas ici.
const UPLOAD_CHUNK_SIZE = 8 * 1024 * 1024;
const UPLOAD_MAX_RETRIES = 3;

async function postToAppsScript(payload) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  return res.json();
}

// Envoie le fichier à l'Apps Script existant, qui l'enregistre dans un dossier
// Google Drive dédié et renvoie son URL de partage. Gratuit : utilise le quota
// Drive du compte Google propriétaire du script, pas de facturation Firebase.
//
// Le fichier est découpé en morceaux de UPLOAD_CHUNK_SIZE envoyés l'un après
// l'autre ; côté Apps Script, chaque morceau est relayé vers une session
// "resumable" de l'API Drive (voir handleChunkUpload dans Code.gs). Cela
// permet d'envoyer des fichiers de plusieurs centaines de Mo, voire plusieurs
// Go, sans jamais dépasser la limite de taille d'une requête Apps Script.
// onProgress(percent) est appelé après chaque morceau si fourni.
// Devine le type MIME à partir de l'extension du fichier lorsque le navigateur
// ne le fournit pas (file.type vide) — évite d'envoyer un mauvais type par
// défaut (ex: "application/pdf" pour un .zip), qui fait échouer l'upload
// côté serveur avec une erreur "Unknown type".
const EXTENSION_MIME_TYPES = {
  pdf: "application/pdf",
  zip: "application/zip",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  txt: "text/plain",
  csv: "text/csv",
  rar: "application/vnd.rar",
  "7z": "application/x-7z-compressed",
};

function guessMimeType(fileName) {
  const ext = (fileName || "").split(".").pop()?.toLowerCase();
  return (ext && EXTENSION_MIME_TYPES[ext]) || null;
}

async function uploadFileToDrive(file, folderId = null, onProgress) {
  const kind = folderId ? "admin" : "piece_jointe";
  const mimeType = file.type || guessMimeType(file.name) || "application/octet-stream";
  const totalSize = file.size;
  const totalChunks = Math.max(1, Math.ceil(totalSize / UPLOAD_CHUNK_SIZE));

  let sessionUri = null;

  for (let i = 0; i < totalChunks; i++) {
    const start = i * UPLOAD_CHUNK_SIZE;
    const end = Math.min(start + UPLOAD_CHUNK_SIZE, totalSize);
    const chunkData = await blobToBase64(file.slice(start, end));

    const payload = {
      type: "upload_chunk",
      kind,
      fileName: file.name,
      mimeType,
      totalSize,
      chunkData,
      rangeStart: start,
      rangeEnd: end - 1,
      sessionUri,
    };
    if (folderId) payload.folderId = folderId;

    let data = null;
    let lastError = null;
    for (let attempt = 1; attempt <= UPLOAD_MAX_RETRIES; attempt++) {
      try {
        data = await postToAppsScript(payload);
        if (!data?.ok) throw new Error(data?.error || "Échec de l'envoi du fichier.");
        lastError = null;
        break;
      } catch (err) {
        lastError = err;
        if (attempt < UPLOAD_MAX_RETRIES) await new Promise((r) => setTimeout(r, 1000 * attempt));
      }
    }
    if (lastError) throw lastError;

    sessionUri = data.sessionUri || sessionUri;
    if (onProgress) onProgress(Math.round((end / totalSize) * 100));

    if (data.done) {
      if (!data.fileUrl) throw new Error("Échec de l'upload.");
      return data.fileUrl;
    }
  }
  throw new Error("Échec de l'upload : le fichier n'a pas été confirmé comme reçu.");
}

// Demande à l'Apps Script de supprimer (déplacer dans la corbeille) un fichier
// Drive à partir de son id. Nécessite que l'Apps Script gère un type
// "delete_drive_file" côté backend (voir note fournie avec ce fichier).
async function deleteFileFromDrive(fileId) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ type: "delete_drive_file", fileId }),
  });
  const data = await res.json();
  if (!data?.ok) throw new Error(data?.error || "Échec de la suppression.");
  return true;
}
import {
  Headset, ShieldCheck, Package, GraduationCap, ChevronLeft,
  FileCheck, RefreshCw, Clock,
  Search, PackageSearch, ClipboardList, BookOpen, Video, Award,
  FileText, Server, Wind, Gauge, Truck, Disc,
  MonitorSmartphone, Cpu, SlidersHorizontal, Thermometer, Droplet,
  Component, PlugZap, Eye, X, ChevronRight, PanelLeft, RotateCcw, RotateCw, Settings,
  ArrowRight, Paperclip, Snowflake, Layers, SeparatorVertical, Filter, Waves, Droplets, ArrowDownToLine, Bell,
  Archive, CheckCircle2, History, Inbox, Download, Send, Trash2, ZoomIn, ZoomOut, Upload
} from "lucide-react";

import logo from "./assets/logo-kaeser.png";
import imagePoleExpert from "./assets/image-pole-expert.png";
import iconSupportTechnique from "./assets/Support.png";
import iconGarantie from "./assets/Garantie.png";
import iconPiecesDetachees from "./assets/Pièces.png";
import iconFormation from "./assets/Formation.png";
import iconInfo from "./assets/Info.png";
import imgTooltipMachine from "./assets/Machine.png";
import imgTooltipNumeroSerie from "./assets/No série.png";
import imgTooltipReference from "./assets/Référence.png";
import imgCompresseur from "./assets/Compresseur.png";
import imgVisSeche from "./assets/Vis sèche.png";
import imgSurpresseur from "./assets/Surpresseur.png";
import imgMobilair from "./assets/Mobilair.png";
import imgPiston from "./assets/Piston.png";
import imgTraitement from "./assets/Traitement.png";
import imgSigmaControl from "./assets/Sigma control.png";
import imgSAM from "./assets/SAM.png";
import imgVariateur from "./assets/Variateur.png";
import imgInstruments from "./assets/Instruments.png";
import imgHuile from "./assets/Huile.png";
import imgAda from "./assets/Ada.png";
import imgBelimo from "./assets/BELIMO.png";
import imgSigmaScb from "./assets/SCB.png";
import imgSigmaSc1 from "./assets/SC1.png";
import imgSigmaSc2 from "./assets/SC2.png";
import imgSigmaSc3 from "./assets/SC3.png";
import imgSigmaScm from "./assets/SCM.png";
import imgSecheurFrigorifique from "./assets/Sécheur frigorifique.png";
import imgSecheurAdsorption from "./assets/Sécheur adsorption.png";
import imgSecheurMembrane from "./assets/Sécheur à membrane.png";
import imgFiltration from "./assets/Filtration.png";
import imgVanneDhs from "./assets/Vanne DHS.png";
import imgAquamat from "./assets/Aquamat.png";
import imgPurgeur from "./assets/Purgeur.png";
import imgSecheurHybritec from "./assets/Hybritec.png";
import imgSecheurCalorsec from "./assets/Calorsec.png";
import imgSchemaTemperatureRD from "./assets/Temperature_RD.png";
import imgSchemaPidEau from "./assets/PID_Eau.png";
import imgSchemaPTChargeVide from "./assets/P-T_Charge-Vide.png";
import imgSchemaPidAir from "./assets/PID_Air.png";
pdfjs.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.js`;

const DRIVE_API_KEY = "AIzaSyBCzsfVrBWSXSxS_5cVi0ESsQ7cqiNXtPg";

async function getDriveFiles(folderId, extension = null) {
  const listUrl = new URL("https://www.googleapis.com/drive/v3/files");
  let q = `'${folderId}' in parents and trashed=false`;
  if (extension) {
    const extensions = Array.isArray(extension) ? extension : [extension];
    const extQuery = extensions.map((ext) => `name contains '.${ext}'`).join(" or ");
    q += ` and (${extQuery})`;
  }
  
  listUrl.searchParams.set("q", q);
  listUrl.searchParams.set("fields", "files(id,name)");
  listUrl.searchParams.set("key", DRIVE_API_KEY);

  const res = await fetch(listUrl);
  if (!res.ok) throw new Error("Erreur API Drive");
  const data = await res.json();

  return (data.files || []).map((f) => ({
    id: f.id,
    name: f.name,
    url: `https://www.googleapis.com/drive/v3/files/${f.id}?alt=media&key=${DRIVE_API_KEY}`,
  })).sort((a, b) => a.name.localeCompare(b.name));
}

async function getDriveDocuments(folderId) {
  return getDriveFiles(folderId, null);
}

// Le script d'upload renvoie un lien de visualisation Drive
// (ex. https://drive.google.com/file/d/ID/view) : parfait pour un lien
// cliquable, mais inutilisable comme src d'une balise <img> (Drive renvoie
// une page HTML, pas les octets de l'image, donc rien ne s'affiche).
// On le convertit à l'affichage en URL directement embarquable, sans
// toucher aux données déjà enregistrées.
function driveEmbeddableUrl(url) {
  if (typeof url !== "string" || !url) return url;
  const match = url.match(/\/file\/d\/([-\w]{20,})/) || url.match(/[?&]id=([-\w]{20,})/);
  const id = match && match[1];
  if (!id) return url; // pas un lien Drive reconnu, on le laisse tel quel
  return `https://drive.google.com/thumbnail?id=${id}&sz=w1000`;
}

// Réécrit les src des <img> dans un fragment HTML enregistré (message ou
// réponse) pour les rendre embarquables, sans modifier le HTML stocké.
function makeImagesEmbeddable(html) {
  if (!html) return html;
  return html.replace(/<img([^>]*?)\ssrc="([^"]*)"/g, (full, attrs, src) => `<img${attrs} src="${driveEmbeddableUrl(src)}"`);
}

// Dossier Drive contenant le PDF "Conditions générales de vente et de garantie"
const GARANTIE_CGV_FOLDER_ID = "1lJ1Lnpj1veKBcqLgbGTZYP8f7-6R_dx_";

// --- Notifications de nouveaux documents PDF ---

const DRIVE_FOLDER_LABELS = {
  driveFolderId: "Notice d'utilisation",
  driveFolderIdCodesDefaut: "Codes défaut",
  driveFolderIdInstructionTechnique: "Instruction technique",
  driveFolderIdCommunication: "Communication",
  driveFolderIdUpdate: "Update",
  driveFolderIdServiceInstruction: "Service instruction",
  driveFolderIdInstructionMontage: "Instruction de montage",
};

function collectDriveFolderTargets(nodes, pathLabels = []) {
  let results = [];
  for (const node of nodes) {
    const newPath = node.label ? [...pathLabels, node.label] : pathLabels;
    for (const [key, docLabel] of Object.entries(DRIVE_FOLDER_LABELS)) {
      if (node[key]) {
        results.push({ folderId: node[key], context: `${newPath.join(" > ")} — ${docLabel}` });
      }
    }
    if (Array.isArray(node.items) && node.items.length) {
      results = results.concat(collectDriveFolderTargets(node.items, newPath));
    }
  }
  return results;
}

const PDF_SCAN_THROTTLE_MS = 15 * 60 * 1000; // ne relance un scan complet que toutes les 15 minutes

async function scanDriveForNewPdfNotifications(force = false) {
  const scanStateRef = doc(db, "appState", "pdfNotificationScan");

  let lastScanAt = 0;
  try {
    const scanStateSnap = await getDoc(scanStateRef);
    lastScanAt = scanStateSnap.exists() ? scanStateSnap.data()?.lastScanAt?.toMillis?.() || 0 : 0;
  } catch (err) {
    console.error("[pdfNotifications] Lecture de l'état du scan impossible — vérifiez les règles Firestore sur la collection 'appState'.", err);
    return { ok: false, reason: "read-scan-state-failed", error: err };
  }

  const now = Date.now();
  if (!force && now - lastScanAt < PDF_SCAN_THROTTLE_MS) {
    console.info(`[pdfNotifications] Scan ignoré : dernier scan il y a ${Math.round((now - lastScanAt) / 1000)}s (throttle 15 min). Utilisez force=true pour forcer.`);
    return { ok: true, skipped: true };
  }

  try {
    await setDoc(scanStateRef, { lastScanAt: serverTimestamp() }, { merge: true });
  } catch (err) {
    console.error("[pdfNotifications] Écriture de l'état du scan impossible — vérifiez les règles Firestore sur la collection 'appState'.", err);
    return { ok: false, reason: "write-scan-state-failed", error: err };
  }

  const targets = collectDriveFolderTargets(CATEGORIES);

  let existingSnap;
  try {
    existingSnap = await getDocs(collection(db, "pdfNotifications"));
  } catch (err) {
    console.error("[pdfNotifications] Lecture des notifications existantes impossible — vérifiez les règles Firestore sur la collection 'pdfNotifications'.", err);
    return { ok: false, reason: "read-notifications-failed", error: err };
  }
  const knownIds = new Set(existingSnap.docs.map((d) => d.id));

  let created = 0;
  let driveErrors = 0;
  let writeErrors = 0;

  for (const target of targets) {
    let files = [];
    try {
      files = await getDriveFiles(target.folderId, "pdf");
    } catch (err) {
      driveErrors++;
      console.warn(`[pdfNotifications] Erreur Drive sur le dossier ${target.folderId} (${target.context}) — vérifiez que le dossier est bien partagé publiquement et que la clé API Drive est valide.`, err);
      continue;
    }
    for (const file of files) {
      if (knownIds.has(file.id)) continue;
      knownIds.add(file.id);
      try {
        await setDoc(doc(db, "pdfNotifications", file.id), {
          fileName: file.name,
          fileUrl: file.url,
          folderId: target.folderId,
          context: target.context,
          createdAt: serverTimestamp(),
          validatedBy: [],
        });
        created++;
      } catch (err) {
        writeErrors++;
        console.error(`[pdfNotifications] Création de la notification impossible pour "${file.name}" — vérifiez les règles Firestore sur 'pdfNotifications'.`, err);
      }
    }
  }

  console.info(`[pdfNotifications] Scan terminé : ${targets.length} dossier(s) analysé(s), ${created} nouvelle(s) notification(s), ${driveErrors} erreur(s) Drive, ${writeErrors} erreur(s) Firestore.`);
  return { ok: true, created, driveErrors, writeErrors, foldersScanned: targets.length };
}

const PDF_MODULES = import.meta.glob("/src/documents/**/*.pdf", {
  eager: true,
  query: "?url",
  import: "default",
});

function getLocalDocuments(itemId) {
  return Object.entries(PDF_MODULES)
    .filter(([path]) => path.includes(`/documents/${itemId}/`))
    .map(([path, url]) => ({
      id: path,
      url,
      name: path.split("/").pop().replace(/\.pdf$/i, ""),
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

const COLORS = {
  navy: "#0B1F3A",
  gold: "#F0AB00",
  goldDark: "#C98F00",
  bannerGray: "#7C8B8D",
  bg: "#FAFAF8",
  cardBorder: "#E7E7E3",
  textMuted: "#6B7280",
};

const SERVICE_EMAILS = {
  "Garantie": "garantie.france@kaeser.com",
  "Pièces détachées": "offres.france@kaeser.com",
};

// Le statut administrateur est désormais déterminé par la présence d'un document
// dans la collection Firestore "admins" (id du document = e-mail de l'utilisateur).
// Voir isAdmin dans App() plus bas, et firestore.rules pour la sécurisation côté serveur.

const CATEGORIES = [
  {
    id: "support-technique",
    label: "Support technique",
    description: "Documentation et assistance",
    icon: Headset,
    image: iconSupportTechnique,
    searchable: true,
    itemIconColor: "#5B7C87",
    items: [
      {
        id: "compresseurs-vis", label: "Compresseurs à vis", icon: Server, image: imgCompresseur,
        items: [
          { id: "sc2-vis", label: "Compresseur SC2", icon: Settings, image: imgCompresseur, driveFolderId: "1QAyqki_IItZ6vOtf2EuADDKmKGbLkFkU",driveFolderIdCodesDefaut: "1mTt31_Kzc17wIWHVsQxay4aPi3IFvOxN",driveFolderIdInstructionTechnique: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3",driveFolderIdServiceInstruction: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3",driveFolderIdInstructionMontage: "12J6_w7WTsIQs8JcMmtnWYxMbPJDA9Y8X",
          driveFolderIdDocuments: "1giNk1A9sRJUpdBuGvY0zbDKSeX6gn1Sh"
          },
          { id: "sc3-vis", label: "Compresseur SC3", icon: Settings, image: imgCompresseur, driveFolderId: "19Z-TuSJgRa3Aq3btywU2AsdT746WJwtX",driveFolderIdCodesDefaut: "1PqOMKKr2GeUnIfp2DscllnaQ_L-NFU43",driveFolderIdInstructionTechnique: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3",driveFolderIdServiceInstruction: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3", driveFolderIdInstructionMontage: "12J6_w7WTsIQs8JcMmtnWYxMbPJDA9Y8X",
          driveFolderIdDocuments: "1giNk1A9sRJUpdBuGvY0zbDKSeX6gn1Sh"
          },
          { id: "scb-vis", label: "Compresseur SCB", icon: Settings, image: imgCompresseur, driveFolderId: "1PqDpXZQGdGXvwpV421WyY9scanMmCQhC",driveFolderIdCodesDefaut: "1Qhv79NL1RaAw-J1oR4uUbTfeqbjZmX7C",driveFolderIdInstructionTechnique: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3",driveFolderIdServiceInstruction: "1z83YLSzx3WrX8hdJ2I52rvwo-KOynPc3", driveFolderIdInstructionMontage: "12J6_w7WTsIQs8JcMmtnWYxMbPJDA9Y8X",
          driveFolderIdDocuments: "1giNk1A9sRJUpdBuGvY0zbDKSeX6gn1Sh"
          },
        ],
      },
      { id: "vis-seche", label: "Vis sèche", icon: Server, image: imgVisSeche, driveFolderId: "1TDsbAxJXJ0k55rM7EbExljTInHqZCtbh", driveFolderIdCodesDefaut: "1YcFi78aP2YlWC8lw6zoki9gTdrfm9DOb" ,driveFolderIdInstructionTechnique: "1TTViFu40NERinMyqJFS2DJPU62giKX31",driveFolderIdServiceInstruction: "10s4eqnK0uvtQHDOvBR3S2dzvAwqPoom5", driveFolderIdInstructionMontage: "1PnUiU6CSE1N_fYuuC2PWvu3Fv0pBzswA", driveFolderIdDocuments: "1BmcWZTKiWYTg8Jh4h3NjLTFO5X-53Nl3",
        schemas: [
          { label: "Température / pression sécheur RD", src: imgSchemaTemperatureRD },
          { label: "PID Vis sèche ref. eau", src: imgSchemaPidEau },
          { label: "Pression/Température en charge et à vide", src: imgSchemaPTChargeVide },
          { label: "PID vis sèche ref. air", src: imgSchemaPidAir },
        ]      
      },


      
      { id: "surpresseur-vis", label: "Surpresseur à vis", icon: Gauge, image: imgSurpresseur, driveFolderId: "1p8bsMe2iw688-mmNW9ea8uPZhrrLxNBV" ,driveFolderIdCodesDefaut: "1y5OppgzpaFjvC7Qq9iMgnamyot1x9h7k" ,driveFolderIdInstructionTechnique: "1u_YkofGRa2q_LS0RrqWPrYIWDIzvxLRd",driveFolderIdServiceInstruction: "1k5b50Pex_F2zS3rTClCpO_Dvdz9oq77o", driveFolderIdInstructionMontage: "1aMbCMW8g2cREDbNrtaMEEYI6GkmqpyUO", driveFolderIdDocuments: "1brzVMZH9u8IRAsBWo0CAFJs-UnND1Zp1"    
      },
      
      { id: "mobilair", label: "Mobilair", icon: Truck, image: imgMobilair, driveFolderId: "18oDG68eFeXA_OV8LrHMFR6nnp4VpWYSO",
  driveFolderIdCodesDefaut: "12cBjWHm8gQ6mQPTwVZ83GgbPx4a0UE4X" ,driveFolderIdInstructionTechnique: "1MDyufp0LGVIJNbgpjDeedK4TLvq2dOtf",driveFolderIdServiceInstruction: "1RIPmBkySMyq2K02JRq9nmAI5-7GNBRbH", driveFolderIdInstructionMontage: "17V1qV0uy_6R4G7T3pmzA-izRlm4pmqNw",
  driveFolderIdDocuments: "1ICSy3pkF3-bR6v6UcKWb3yxV3CpKdRLU"
      },
 
      { id: "piston", label: "Piston", icon: Disc, image: imgPiston, driveFolderId: "1UG8Gd2Lb0682uhR1EltjOAU9tkhhjZMf", driveFolderIdCodesDefaut: "1_hhjsl7E6PT3mDCJXVatfG7aH-SFaPF-", driveFolderIdInstructionTechnique: "1bE65kkKKOElA86jFm82P7DDIaXI3-tpj",driveFolderIdServiceInstruction: "1ZCMgH9iDGmb5FVKxxxVBtIsmEMPjXS0C", driveFolderIdInstructionMontage: "13aOt4_-F8cEzSaKDKRtrMFuS9ZcliumN", driveFolderIdDocuments: "1urB9RdhQoZTWO5qoaPZlRkhOgaZU7x4T" 
      },
      
      {
        id: "traitement-air", label: "Traitement d'air", icon: Wind, image: imgTraitement,
        items: [
          { id: "secheur-frigorifique", label: "Sécheur frigorifique", icon: Snowflake, image: imgSecheurFrigorifique,
          driveFolderId: "1DqiAzIuA7Tt37ZS0j-RqRu7fTdPLb11u",driveFolderIdCodesDefaut: "1-YSG7A-7eMCN8y-_5y_AvlezIJoEA-gr",driveFolderIdUpdate: "1pTe9MUDxUlcFNZ1Qo6Yo-Mua5Pce4LQx",driveFolderIdCommunication: "1LtNXgTccs0rEOPv89FJI_UkgAcV6uakw",driveFolderIdInstructionTechnique: "1ehJUSnuC3FIF2MtcdHWp_QaZg6DtSxPd",driveFolderIdServiceInstruction: "19WR5jrwLFO8pJUwB5gc2tmih1CEfhrIn", driveFolderIdInstructionMontage: "1zRor9rVeZHPk7mEFEpSSBhiHncvUxh2A"    
          },
          
          { id: "secheur-adsorption", label: "Sécheur adsorption", icon: Layers, image: imgSecheurAdsorption,
 driveFolderId: "1EEXnRK2rrG0P8p9Y9r4CL28LggEwytel",driveFolderIdUpdate: "1cgsVSfxvfp1D9LJqMI3rGvt23VU6lIfe",driveFolderIdCommunication: "14N4pqCXLQO5Lm55dEnCq2IIF46azUlDA",driveFolderIdInstructionTechnique: "1r35agCwTrMvD5UlE5L9y_TUAFEJDHxGF",driveFolderIdServiceInstruction: "187ZGxeUtv9y9KHASJC8Ziz0SFLiboT2X", driveFolderIdInstructionMontage: "18Hx0s23OUz0gdfWaB551xtutHgG5pSZ_"     
          },
          
          { id: "secheur-Calorsec", label: "Sécheur Calorsec", icon: Layers, image: imgSecheurCalorsec,driveFolderId: "1hrLiliSmWXYL0FLQx-EPcR8QpGkV-lNw",driveFolderIdInstructionTechnique: "11ok1zv1qAXhSF8dpRrZsz7K6VemwXZ1L",driveFolderIdServiceInstruction: "1G1HYEAD3BItkLEJxnpyC29ry_BgWNAmc", driveFolderIdInstructionMontage: "1aMbCMW8g2cREDbNrtaMEEYI6GkmqpyUO"     
          },
              
          { id: "secheur-Hybritec", label: "Sécheur Hybritec", icon: Layers, image: imgSecheurHybritec,driveFolderId: "1Tcdva_KsgIREt_hdhTA6AUY0h9uoNmm2",driveFolderIdInstructionTechnique: "1wpOx46onH3j0cDYxcPJ-VlaSHwSoaRm3",driveFolderIdServiceInstruction: "1h8evefbqvUs3Rv9TRZLjxp4IadMeUkEe", driveFolderIdInstructionMontage: "1s9FlSgFwvRI2V57fPcOMolS_nRLG7uRT"
          },
          
          { id: "secheur-membrane", label: "Sécheur à membrane", icon: SeparatorVertical, image: imgSecheurMembrane,driveFolderId: "1ZBp2QUqSn3WTaOIUs-XLmJxE6XuVc2Dw",driveFolderIdInstructionTechnique:"1aQrCcUS9MZznIqSTAoUd0HJXO1jH7hx-",driveFolderIdServiceInstruction: "1xbJyZpPWn37Sx7kXzbumoS9B-V0RX3EF", driveFolderIdInstructionMontage: "1XEwCnFivrSY1dpWAZQYHeZVMMWASNkt9"
          },
          
          { id: "filtration", label: "Filtration", icon: Filter, image: imgFiltration,driveFolderId: "1h8_lbyHTGtgyAuCg2hJeWjXx3Q5mUBOW",driveFolderIdInstructionTechnique:"1uY2qHgSGgV_XP9Hx3-z-le2o9vRzTr7t",driveFolderIdServiceInstruction: "16TZ4iZVI_6ChTCGEelY7MFPB6IrpvdUy",driveFolderIdServiceInstruction: "16TZ4iZVI_6ChTCGEelY7MFPB6IrpvdUy", driveFolderIdInstructionMontage: "1Jqp1EHnhHzesiEyc2MfGCNmpwtN-aIbQ"
          
          
          },
          { id: "vanne-dhs", label: "Vanne DHS", icon: Waves, image: imgVanneDhs,driveFolderId: "1nnJz39tLbHOIM8IsKyk-64XH4WmRSX6k",driveFolderIdUpdate: "1aQc-YqYy-AyKjB0uZiwa8AAKuaW4nG2G",driveFolderIdCommunication: "1TpuyK0F43JWDdQg0R8WE63anKa4QVDNX",driveFolderIdInstructionTechnique:"1C5djRMeStx4ZoUSMyojiMfV2nBWiwy7P",driveFolderIdServiceInstruction: "1AeM9f-ssDIt-T1_m1416hnYcxWAg2Owq", driveFolderIdInstructionMontage: "1gTaa-7u2b6OmG7b3BhzkALxYdZvFQsXZ"
          },
          
          { id: "aquamat", label: "AQUAMAT", icon: Droplets, image: imgAquamat,driveFolderId: "1Aa9AGQJGpx5zWieNKk1J1QXdM9-D09Wp",driveFolderIdCodesDefaut: "11iXVrjDu_bkA6bKH_srWm2JK6JVEoTeK",driveFolderIdInstructionTechnique:"1ahJBjKap0MJ5HLNogfOESpUS3eDAMKcZ",driveFolderIdServiceInstruction: "16TZ4iZVI_6ChTCGEelY7MFPB6IrpvdUy",driveFolderIdServiceInstruction: "1tI_cJdm25oHyHCaPQMHaIwPm359SPcGk", driveFolderIdInstructionMontage: "14Kzy5-91uEt2g-cyQUmKr82fn6xj8PVq"
          },
          
          { id: "purgeur", label: "Purgeur", icon: ArrowDownToLine, image: imgPurgeur,driveFolderId: "1PyTxu_ZoNfaDXJR9OeJo_YVohPyO9-Vj",driveFolderIdInstructionTechnique:"1NpPgubPIV6dLN1Qv-vAzgozk_X-pf-y4",driveFolderIdServiceInstruction: "16TZ4iZVI_6ChTCGEelY7MFPB6IrpvdUy",driveFolderIdServiceInstruction: "1ScDvRBnwXHBFX3hBmVPKyF1NHrLrwWCw", driveFolderIdInstructionMontage: "1ao5KgPXwolYtF5nqnn_1y_RyPALl7BEy"
          },

        { id: "instruments", label: "Instruments", icon: Thermometer, image: imgInstruments,driveFolderId: "1Hjuwji4E8ezY1HHAU7LYwNsJu0AYJFGf",driveFolderIdInstructionTechnique:"1dIgf3cx_mECwjPJCcrC81nR6MwxMUhBB",driveFolderIdServiceInstruction: "1NwhnTNzLdW5kKs351FHnLnyX9srKSQuw",driveFolderIdInstructionMontage: "1d25F_XIMg0CxN21nUHf6jI76Far-lvD_"
      },
      
        ],
      },
      {
        id: "sigma-control", label: "Sigma Control", icon: MonitorSmartphone, image: imgSigmaControl,
        items: [
          { id: "sigma-scb", label: "SIGMA CONTROL BASIC", icon: MonitorSmartphone, image: imgSigmaScb, driveFolderIdCodesDefaut: "1Qhv79NL1RaAw-J1oR4uUbTfeqbjZmX7C", driveFolderIdInstructionTechnique: "1lrCFIdbOxnHZo5Mdtd7Ycs6k0xW-Zprj" },
          { id: "sigma-sc1", label: "SIGMA CONTROL 1", icon: MonitorSmartphone, image: imgSigmaSc1, driveFolderId: "1_-Y7puS8dMZDyAzzQwJkAHrOxvONsx50", driveFolderIdCommunication: "18y9gRyp4i12I6MUA32sbcf7qzL9eyreY" },
          { id: "sigma-sc2", label: "SIGMA CONTROL 2", icon: MonitorSmartphone, image: imgSigmaSc2, driveFolderId: "179_lD8DVt5RHBMxCEk28Szyeyz2iQ3qw", driveFolderIdCodesDefaut: "1mTt31_Kzc17wIWHVsQxay4aPi3IFvOxN", driveFolderIdCommunication: "1QjspPJ9dln1GzAMjzU0lK_9CrK82Sptn", driveFolderIdUpdate: "1hyQeeah1J_6qdSBz2YmI6qKkpGe7jGNw", driveFolderIdInstructionTechnique: "1VVNszF7ZH3bmjtGdZRfdmB20SHwFKUCa", simulateurUrl: "https://i0070916-p50100-c1-hc3fe5i5e5rdxznzctpjng5mx4sqx66dr.webdirect.mdex.de/index.html", simulateurLabel: "Simulateur SC 2" },
          { id: "sigma-sc3", label: "SIGMA CONTROL 3", icon: MonitorSmartphone, image: imgSigmaSc3, driveFolderId: "1-idZVZsifWbGz3wViU0rABYiKTqkN6Lr", driveFolderIdCodesDefaut: "1PqOMKKr2GeUnIfp2DscllnaQ_L-NFU43", driveFolderIdCommunication: "1P2SxMX7nIoZoT2B1RpqsykVMoKORHffN", driveFolderIdInstructionTechnique: "1wQeZ3YZdmfTCyZ2bCMZC1Tj-o9ur0lzW", driveFolderIdUpdate: "1CtId3no7x1whbTIef05hE4y1LnvBSFgZ", simulateurUrl: "https://i0190606-guacamole.direct.mdex.de/guacamole/#/", simulateurLabel: "Simulateur SC 3", simulateurCredentials: { username: "KAESER-SC3", password: "SigmaControl3" } },
          { id: "sigma-scm", label: "SIGMA CONTROL MOBILE", icon: MonitorSmartphone, image: imgSigmaScm, driveFolderId: "1lIhU9S5N4cSX5RBecrdROMwn7tHagZIi", driveFolderIdCodesDefaut: "12cBjWHm8gQ6mQPTwVZ83GgbPx4a0UE4X", driveFolderIdUpdate: "1Xh_sSns43pWKGi09Bf6hoae2vAyBuwzA", driveFolderIdInstructionTechnique: "1re_FkgT4jFWvPpZ97X8mcwqidJ6uGD-s" },
        ],
      },
      { id: "sam-40", label: "SAM 4.0", icon: Cpu, image: imgSAM, driveFolderId: "1Aq-14EXHEXNX35RUfXHY_cXEsmjJTrGX",driveFolderIdCodesDefaut: "1qXwsn1pfuvIJOE9JMjm1tLFMI4b55Pnv",driveFolderIdUpdate: "1RHhF45uJrn4VLVUsCjnjVTdCWdJEK6jX",driveFolderIdCommunication: "1cbz_6HQzrAGVmH-llSKwaljOBhTAC8xO",driveFolderIdInstructionTechnique: "1JtRvY17NDOe_nazvceeOPYRR8BpuYTcM",driveFolderIdServiceInstruction: "1cQmyE9YnCdrUAkp26_7hFkIQnmYk9sV2", driveFolderIdInstructionMontage: "1xYAXNyoqDGzNtNUM7SGMj-CLJ1I0Lwm9"    
      
      
      },
      { id: "variateur", label: "Variateur", icon: SlidersHorizontal, image: imgVariateur,driveFolderId: "1C9I8HQCRiwkLC9ABb_n_BhVl4KmXarxy",driveFolderIdCodesDefaut: "194uOKt8oOGrYZ6pgKpC5TdAIzDrrrEd1",driveFolderIdInstructionTechnique: "1Un3FfzsZBSqLJndSJ-n5qbLGvnezLRqS",driveFolderIdServiceInstruction: "1xG5Kl7zkljhduCY78FlUqd1luuvM3jHN", driveFolderIdInstructionMontage: "1NQ7q2HayMobzMZEV3o50PJpwOYAM6ABZ", driveFolderIdDocuments: "1xKZikqpvSLuzfzl1BD5axnnw6BhYR0cP"      
      },
           
      { id: "huile", label: "Huile", icon: Droplet, image: imgHuile, driveFolderIdInstructionTechnique: "1vD-kDI5DyKA6PQHMqgWtUgAJPo62Nfsu" },
      { id: "ada", label: "ADA", icon: Component, image: imgAda, driveFolderIdInstructionTechnique: "1UCzv0J4YNOG9Wvyu2GJFK1xllMK9TJ6n" },
      { id: "belimo", label: "BELIMO", icon: PlugZap, image: imgBelimo, driveFolderIdInstructionTechnique: "10NZSFUUzsV6fR5aY7mWqVivQDMaJjP63" },
    ],
  },
  {
    id: "garantie",
    label: "Garantie",
    description: "Enregistrements et demandes",
    icon: ShieldCheck,
    image: iconGarantie,
    driveFolderId: GARANTIE_CGV_FOLDER_ID,
    items: [
      { id: "demande-garantie", label: "Demande garantie", icon: FileCheck },
      { id: "enregistrement-machine", label: "Enregistrement de la machine", icon: ClipboardList },
    ],
  },
  {
    id: "pieces-detachees",
    label: "Pièces détachées",
    description: "Catalogues et références",
    icon: Package,
    image: iconPiecesDetachees,
    // Dossier Drive contenant les catalogues de pièces détachées.
    driveFolderId: "1q0aUSfqOnNwjaKqDq0wsKN9Eq8C4ych7",
    items: [],
  },
  {
    id: "formation",
    label: "Formation",
    description: "Modules experts",
    icon: GraduationCap,
    image: iconFormation,
    driveFolderId: "1jdNOxzb3yREl6X4yNpyxZGvhasujbymC",
    items: [],
  },
];

// Détecte les écrans étroits (mobile/Android) et se met à jour au redimensionnement/rotation.
function useIsMobile(breakpoint = 640) {
  const [isMobile, setIsMobile] = useState(
    () => typeof window !== "undefined" && window.innerWidth <= breakpoint
  );
  useEffect(() => {
    const mql = window.matchMedia(`(max-width: ${breakpoint}px)`);
    const handler = (e) => setIsMobile(e.matches);
    setIsMobile(mql.matches);
    mql.addEventListener ? mql.addEventListener("change", handler) : mql.addListener(handler);
    return () => {
      mql.removeEventListener ? mql.removeEventListener("change", handler) : mql.removeListener(handler);
    };
  }, [breakpoint]);
  return isMobile;
}

export default function App() {
  const [user, setUser] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [kickedOut, setKickedOut] = useState(false);
  const [pdfNotifications, setPdfNotifications] = useState([]);
  const [showNotifModal, setShowNotifModal] = useState(false);
  const [scanningNotifs, setScanningNotifs] = useState(false);
  const autoOpenedNotifRef = useRef(false);
  const [isFirstLogin, setIsFirstLogin] = useState(false);
  const firstLoginValidatedRef = useRef(false);
  const [expertRequests, setExpertRequests] = useState([]);
  const [showExpertRequestsModal, setShowExpertRequestsModal] = useState(false);
  const [showAdminUploadModal, setShowAdminUploadModal] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthReady(true);
      if (u) setKickedOut(false);
    });
    return unsubscribe;
  }, []);

  useEffect(() => {
    if (!user?.email) { setIsAdmin(false); return; }
    let cancelled = false;
    getDoc(doc(db, "admins", user.email))
      .then((snap) => { if (!cancelled) setIsAdmin(snap.exists()); })
      .catch(() => { if (!cancelled) setIsAdmin(false); });
    return () => { cancelled = true; };
  }, [user]);

  useEffect(() => {
    if (!user) return;
    scanDriveForNewPdfNotifications();
    // Relance le scan périodiquement et quand l'utilisateur revient sur l'onglet,
    // pour détecter les PDF ajoutés pendant que l'app est déjà ouverte.
    const interval = setInterval(() => scanDriveForNewPdfNotifications(), 5 * 60 * 1000);
    const onVisibilityChange = () => { if (document.visibilityState === "visible") scanDriveForNewPdfNotifications(); };
    document.addEventListener("visibilitychange", onVisibilityChange);
    const unsubscribe = onSnapshot(collection(db, "pdfNotifications"), (snap) => {
      const list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      list.sort((a, b) => (b.createdAt?.toMillis?.() || 0) - (a.createdAt?.toMillis?.() || 0));
      setPdfNotifications(list);
    });
    return () => {
      clearInterval(interval);
      document.removeEventListener("visibilitychange", onVisibilityChange);
      unsubscribe();
    };
  }, [user]);

  const handleForceScan = async () => {
    setScanningNotifs(true);
    const result = await scanDriveForNewPdfNotifications(true);
    setScanningNotifs(false);
    return result;
  };

  useEffect(() => {
    if (!user) return;
    // Pour un utilisateur non-admin, on interroge directement ses propres demandes (where
    // requestedBy == son e-mail) plutôt que de lire toute la collection puis filtrer côté
    // client : si les règles Firestore restreignent la lecture aux demandes dont on est
    // l'auteur (ou aux admins), un onSnapshot sans "where" correspondant échoue silencieusement
    // pour un non-admin (permission-denied, sans gestionnaire d'erreur), ce qui vidait à tort
    // "En cours" ET "Historique" même quand l'utilisateur avait bien des demandes.
    const expertRequestsQuery = isAdmin
      ? collection(db, "expertRequests")
      : query(collection(db, "expertRequests"), where("requestedBy", "==", user.email));
    const unsubscribe = onSnapshot(
      expertRequestsQuery,
      (snap) => {
        const list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        list.sort((a, b) => (b.createdAt?.toMillis?.() || 0) - (a.createdAt?.toMillis?.() || 0));
        setExpertRequests(list);
      },
      (err) => {
        console.error("Erreur d'écoute expertRequests (droits Firestore ?) :", err);
      }
    );
    return unsubscribe;
  }, [user, isAdmin]);

  const visibleExpertRequests = expertRequests;
  const pendingExpertRequestsCount = useMemo(
    () => visibleExpertRequests.filter((r) => !r.archived).length,
    [visibleExpertRequests]
  );

  const pendingNotifications = useMemo(
    () => pdfNotifications.filter((n) => !(n.validatedBy || []).includes(user?.email)),
    [pdfNotifications, user]
  );

  // À la toute première connexion d'un utilisateur (jamais connecté auparavant, quel que
  // soit l'appareil), on ne considère pas les documents déjà présents comme "nouveaux" :
  // ils sont marqués comme vus silencieusement, sans ouvrir la fenêtre de notification.
  // L'information est stockée côté serveur (Firestore), rattachée au compte utilisateur.
  useEffect(() => {
    if (!user) return;
    let cancelled = false;
    const userMetaRef = doc(db, "users", user.uid);
    getDoc(userMetaRef)
      .then((snap) => {
        if (cancelled) return;
        const alreadyHandled = snap.exists() && snap.data()?.pdfFirstLoginHandled;
        if (!alreadyHandled) {
          setIsFirstLogin(true);
          autoOpenedNotifRef.current = true; // empêche l'ouverture automatique du popup
          setDoc(userMetaRef, { pdfFirstLoginHandled: true }, { merge: true }).catch(() => {});
        }
      })
      .catch(() => {
        // En cas d'erreur de lecture, on n'active pas le mode "première connexion" :
        // le comportement normal (avec ouverture automatique si besoin) s'applique.
      });
    return () => { cancelled = true; };
  }, [user]);

  useEffect(() => {
    if (!isFirstLogin || firstLoginValidatedRef.current || !user || pdfNotifications.length === 0) return;
    firstLoginValidatedRef.current = true;
    Promise.all(
      pdfNotifications
        .filter((n) => !(n.validatedBy || []).includes(user.email))
        .map((n) => updateDoc(doc(db, "pdfNotifications", n.id), { validatedBy: arrayUnion(user.email) }))
    ).catch(() => {});
  }, [isFirstLogin, pdfNotifications, user]);

  useEffect(() => {
    if (!autoOpenedNotifRef.current && pendingNotifications.length > 0) {
      setShowNotifModal(true);
      autoOpenedNotifRef.current = true;
    }
  }, [pendingNotifications.length]);

  useEffect(() => {
    if (!user) return;
    const sessionKey = `pole-expert-session-${user.uid}`;
    let sessionId = localStorage.getItem(sessionKey);
    if (!sessionId) {
      sessionId = crypto.randomUUID();
      localStorage.setItem(sessionKey, sessionId);
    }
    const sessionRef = doc(db, "sessions", user.uid);
    setDoc(sessionRef, { sessionId, updatedAt: serverTimestamp() }).catch(() => {});
    const unsubscribe = onSnapshot(sessionRef, (snap) => {
      const activeSessionId = snap.data()?.sessionId;
      if (activeSessionId && activeSessionId !== sessionId) {
        setKickedOut(true);
        signOut(auth);
      }
    });
    return unsubscribe;
  }, [user]);

  const [stack, setStack] = useState([]);
  const push = (entry) => setStack((s) => [...s, entry]);
  const pop = () => setStack((s) => s.slice(0, -1));
  const goHome = () => setStack([]);
  const isMobile = useIsMobile();

  const current = stack[stack.length - 1];
  const activeCategory = stack.find((s) => s.type === "category")?.data;

  if (!authReady) return null;
  if (!user) {
    return <LoginScreen kickedOut={kickedOut} />;
  }

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, color: COLORS.navy, fontFamily: "'Inter', system-ui, sans-serif", display: "flex", flexDirection: "column", overflowX: "hidden" }}>
      <div style={{ background: "#FFC800", padding: isMobile ? "10px 12px" : "6px 24px 6px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: isMobile ? "8px" : "20px", minHeight: "72px", flexWrap: isMobile ? "wrap" : "nowrap", boxSizing: "border-box", position: "relative" }}>
        <img src={logo} alt="Kaeser Compresseurs" style={{ height: isMobile ? "36px" : "100%", width: "auto", display: "block", background: "transparent", flexShrink: 0, order: 1 }} />
        <div style={{ display: "flex", alignItems: "center", gap: isMobile ? "6px" : "10px", flexShrink: 0, flexWrap: "wrap", justifyContent: "flex-end", order: 2 }}>
          <button onClick={() => setShowNotifModal(true)} aria-label="Notifications" style={{ position: "relative", display: "flex", alignItems: "center", justifyContent: "center", width: "38px", height: "38px", background: "#FFFFFF", border: "1px solid rgba(0,0,0,0.15)", borderRadius: "8px", cursor: "pointer", color: COLORS.navy }}>
            <Bell size={18} />
            {pendingNotifications.length > 0 && (
              <span style={{ position: "absolute", top: "-6px", right: "-6px", minWidth: "18px", height: "18px", padding: "0 4px", borderRadius: "9px", background: "#C0392B", color: "#FFFFFF", fontSize: "11px", fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>{pendingNotifications.length}</span>
            )}
          </button>
          <button onClick={() => setShowExpertRequestsModal(true)} aria-label="Demandes experts" title="Demandes experts" style={{ position: "relative", display: "flex", alignItems: "center", justifyContent: "center", width: "38px", height: "38px", background: "#FFFFFF", border: "1px solid rgba(0,0,0,0.15)", borderRadius: "8px", cursor: "pointer", color: COLORS.navy }}>
            <ClipboardList size={18} />
            {pendingExpertRequestsCount > 0 && (
              <span style={{ position: "absolute", top: "-6px", right: "-6px", minWidth: "18px", height: "18px", padding: "0 4px", borderRadius: "9px", background: "#C0392B", color: "#FFFFFF", fontSize: "11px", fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>{pendingExpertRequestsCount}</span>
            )}
          </button>
          {isAdmin && (
            <button onClick={() => setShowAdminUploadModal(true)} aria-label="Ajouter / Supprimer un document" title="Ajouter / Supprimer un document (admin)" style={{ display: "flex", alignItems: "center", justifyContent: "center", width: "38px", height: "38px", background: "#FFFFFF", border: "1px solid rgba(0,0,0,0.15)", borderRadius: "8px", cursor: "pointer", color: COLORS.navy }}>
              <Upload size={18} />
            </button>
          )}
          <button onClick={() => signOut(auth)} style={{ fontSize: "13px", color: COLORS.navy, background: "#FFFFFF", border: "1px solid rgba(0,0,0,0.15)", borderRadius: "8px", padding: "8px 14px", cursor: "pointer", whiteSpace: "nowrap" }}>
            {isMobile ? "Déconnexion" : `Se déconnecter (${user.email})`}
          </button>
        </div>
        <h1
          style={
            isMobile
              ? { fontSize: "26px", fontWeight: 800, margin: 0, textAlign: "center", color: "#FFFFFF", whiteSpace: "nowrap", order: 3, flexBasis: "100%" }
              : { fontSize: "48px", fontWeight: 800, margin: 0, textAlign: "center", color: "#FFFFFF", position: "absolute", left: "50%", top: "50%", transform: "translate(-50%, -50%)", whiteSpace: "nowrap" }
          }
        >
          Le Pôle Expert
        </h1>
      </div>
      <section style={{ background: COLORS.bannerGray, color: "#FFFFFF", padding: isMobile ? "0 12px" : "0 24px 0 24px", display: "flex", flexDirection: "column", justifyContent: "flex-start" }}>
        <p style={{ fontStyle: "italic", fontSize: "14px", opacity: 0.9, margin: "10px 0 0", textAlign: "center" }}>Qualité, Performance et Satisfaction Client</p>
        <img src={imagePoleExpert} alt="Illustration Pôle Expert" style={{ display: "block", margin: "6px auto 0", height: isMobile ? "140px" : "250px", width: "auto", maxWidth: "100%" }} />
        <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: "8px 10px", marginTop: "14px", fontSize: "13.5px", opacity: 0.85, visibility: stack.length > 0 ? "visible" : "hidden" }}>
          <span style={{ cursor: "pointer" }} onClick={goHome}>Accueil</span>
          {stack.map((s, i) => {
            const isLast = i === stack.length - 1;
            return (
              <span key={i} style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: "8px 10px" }}>
                <span style={{ opacity: 0.6 }}>/</span>
                <span style={{ cursor: isLast ? "default" : "pointer", opacity: 1, background: isLast ? COLORS.gold : "transparent", color: isLast ? COLORS.navy : "inherit", borderRadius: isLast ? "8px" : 0, padding: isLast ? "2px 10px" : 0 }} onClick={() => { if (!isLast) { setStack((prev) => prev.slice(0, i + 1)); } }}>{s.data.label}</span>
              </span>
            );
          })}
        </div>
      </section>
      <main style={{ flex: 1, padding: isMobile ? "12px 12px 32px" : "12px 24px 40px" }}>
        {!current && <MainMenu onSelect={(cat) => push(cat.items ? { type: "category", data: cat } : { type: "item", data: cat })} />}
        {(current?.type === "category" || current?.type === "subcategory") && <SubMenu category={current.data} onSelect={(item) => push(item.items ? { type: "subcategory", data: item } : { type: "item", data: item })} />}
        {current?.type === "item" && <DetailPage category={stack[stack.length - 2]?.data} item={current.data} />}
      </main>
      {showNotifModal && (
        <PdfNotificationsModal
          notifications={pendingNotifications}
          currentUserEmail={user.email}
          onClose={() => setShowNotifModal(false)}
          onForceScan={handleForceScan}
          scanning={scanningNotifs}
        />
      )}
      {showExpertRequestsModal && (
        <ExpertRequestsModal
          requests={visibleExpertRequests}
          isAdmin={isAdmin}
          currentUserEmail={user.email}
          onClose={() => setShowExpertRequestsModal(false)}
        />
      )}
      {isAdmin && showAdminUploadModal && (
        <AdminUploadModal categories={CATEGORIES} onClose={() => setShowAdminUploadModal(false)} />
      )}
    </div>
  );
}

function MainMenu({ onSelect }) {
  return (
    <div style={gridStyle}>
      {CATEGORIES.map((cat) => (
        <Card key={cat.id} icon={cat.icon} image={cat.image} label={cat.label} description={cat.description} onClick={() => onSelect(cat)} />
      ))}
    </div>
  );
}

function SubMenu({ category, onSelect }) {
  const [query, setQuery] = useState("");
  const [showExpertForm, setShowExpertForm] = useState(false);
  const [showMachineForm, setShowMachineForm] = useState(false);
  const isMobile = useIsMobile();
  const filteredItems = category.items.filter((item) => item.label.toLowerCase().includes(query.toLowerCase()));
  const showSideIcon = category.id === "garantie" || category.id === "pieces-detachees" || category.id === "formation";

  const handleItemClick = (item) => {
    if (category.id === "garantie" && item.id === "demande-garantie") { setShowExpertForm(true); return; }
    if (category.id === "garantie" && item.id === "enregistrement-machine") { setShowMachineForm(true); return; }
    onSelect(item);
  };

  const cardsGrid = (
    <div style={showSideIcon ? { display: "flex", flexWrap: "wrap", gap: "18px", justifyContent: category.id === "garantie" ? "flex-start" : "center" } : gridStyle}>
      {filteredItems.map((item) => (
        <Card 
          key={item.id} 
          icon={item.icon} 
          image={item.image} 
          label={item.label} 
          iconColor={category.itemIconColor} 
          onClick={() => handleItemClick(item)} 
          imageSize={item.id === "sigma-scm" ? 100 : undefined}
        />
      ))}
    </div>
  );

  const content = (
    <>
      {category.searchable && (
        <div style={searchWrapStyle}>
          <Search size={18} color={COLORS.textMuted} />
          <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Rechercher..." style={searchInputStyle} />
        </div>
      )}
      {category.id === "formation" || category.id === "pieces-detachees" ? (
        <div style={{ display: "flex", flexDirection: "column", gap: "22px" }}>
          {category.id === "formation" && <FormationDocumentList driveFolderId={category.driveFolderId} categorie={category.label} />}
          <button onClick={() => setShowExpertForm(true)} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "8px", background: COLORS.gold, color: COLORS.navy, border: "none", borderRadius: "10px", padding: "14px 20px", fontSize: "13px", fontWeight: 700, letterSpacing: "0.02em", textTransform: "uppercase", cursor: "pointer", maxWidth: "480px" }}>
            {category.id === "formation" ? "CONTACTER LE SERVICE FORMATION" : "CONTACTER LE SERVICE PIECES"}
            <ArrowRight size={16} />
          </button>
        </div>
      ) : (
        <>
          {cardsGrid}
          {category.searchable && filteredItems.length === 0 && <p style={{ color: COLORS.textMuted, fontSize: "14px" }}>Aucun résultat pour « {query} ».</p>}
          {category.id === "garantie" && (
            <div style={{ marginTop: "22px" }}>
              <GarantieCgvButton categorie={category.label} />
            </div>
          )}
          {category.id === "garantie" && (
            <button onClick={() => setShowExpertForm(true)} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "8px", background: COLORS.gold, color: COLORS.navy, border: "none", borderRadius: "10px", padding: "14px 20px", fontSize: "13px", fontWeight: 700, letterSpacing: "0.02em", textTransform: "uppercase", cursor: "pointer", maxWidth: "480px", marginTop: "16px" }}>
              CONTACTER LE SERVICE GARANTIE
              <ArrowRight size={16} />
            </button>
          )}
        </>
      )}
    </>
  );

  return (
    <div>
      {showSideIcon ? (
        <div style={{ display: "flex", gap: isMobile ? "20px" : "60px", flexDirection: isMobile ? "column" : "row", flexWrap: "wrap", alignItems: isMobile ? "center" : "flex-start" }}>
          <div style={{ flex: isMobile ? "0 0 auto" : "0 0 200px", textAlign: "center" }}>
            {category.image ? (
              <img src={category.image} alt="" style={{ width: 160, height: 160, objectFit: "contain", display: "block", margin: "0 auto" }} />
            ) : (
              <div style={{ display: "inline-flex", padding: "16px", borderRadius: "14px", background: "#FBF1D9" }}>
                <category.icon size={30} color={COLORS.gold} />
              </div>
            )}
          </div>
          <div style={{ flex: 1, minWidth: 0, width: "100%" }}>{content}</div>
        </div>
      ) : (
        content
      )}
      {showExpertForm && <ExpertRequestForm item={null} category={category} initialSujet={category.id === "formation" ? "Formation" : category.id === "pieces-detachees" ? "Pièces détachées" : "Garantie"} onClose={() => setShowExpertForm(false)} />}
      {showMachineForm && <MachineRegistrationForm category={category} onClose={() => setShowMachineForm(false)} />}
    </div>
  );
}

function SimulateurButton({ item, onOpen }) {
  if (!item.simulateurUrl) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
      <button onClick={onOpen} style={{ display: "flex", alignItems: "center", gap: "8px", background: COLORS.navy, color: "#FFFFFF", border: "none", borderRadius: "10px", padding: "14px 20px", fontSize: "13px", fontWeight: 700, letterSpacing: "0.02em", textTransform: "uppercase", cursor: "pointer", whiteSpace: "nowrap" }}>
        {item.simulateurLabel || "Simulateur"}
        <MonitorSmartphone size={16} />
      </button>
      {item.simulateurCredentials && (
        <div style={{ fontSize: "12.5px", color: COLORS.textMuted, lineHeight: 1.6 }}>
          <div>Username: {item.simulateurCredentials.username}</div>
          <div>Password: {item.simulateurCredentials.password}</div>
        </div>
      )}
    </div>
  );
}

function DetailPage({ category, item }) {
  const [showExpertForm, setShowExpertForm] = useState(false);
  const [showSimulator, setShowSimulator] = useState(false);
  const [selectedSchemaId, setSelectedSchemaId] = useState("");
  const isMobile = useIsMobile();
  const simulateurInNoticeRow = Boolean(item.simulateurUrl && item.driveFolderId);
  const hasSchemas = Array.isArray(item.schemas) && item.schemas.length > 0;
  // Réinitialise le schéma affiché quand on change de produit, pour ne pas
  // garder affiché le schéma d'un produit précédent.
  useEffect(() => { setSelectedSchemaId(""); }, [item.id]);
  const selectedSchema = hasSchemas ? item.schemas.find((s) => s.label === selectedSchemaId) : null;
  return (
    <div style={{ display: "flex", columnGap: isMobile ? "20px" : "60px", rowGap: isMobile ? "20px" : "16px", flexDirection: isMobile ? "column" : "row", flexWrap: "wrap", alignItems: isMobile ? "center" : "flex-start", maxWidth: hasSchemas ? "100%" : "1200px" }}>
      <div style={{ flex: isMobile ? "0 0 auto" : "0 0 300px", display: "flex", flexDirection: "column", alignItems: "center" }}>
        {item.image ? (
          <img src={item.image} alt="" style={{ width: 200, height: 200, objectFit: "contain", marginBottom: "10px" }} />
        ) : (
          <div style={{ display: "inline-flex", padding: "16px", borderRadius: "14px", background: "#FBF1D9", marginBottom: "10px" }}>
            <item.icon size={30} color={COLORS.gold} />
          </div>
        )}
        {simulateurInNoticeRow && (
          <div style={{ marginTop: "8px" }}>
            <SimulateurButton item={item} onOpen={() => setShowSimulator(true)} />
          </div>
        )}
        {hasSchemas && (
          <div style={{ marginTop: "18px", width: "100%" }}>
            <h3 style={{ fontSize: "12px", textTransform: "uppercase", letterSpacing: "0.06em", color: COLORS.textMuted, marginBottom: "10px" }}>Schémas techniques</h3>
            <CustomSelect
              value={selectedSchemaId}
              onChange={setSelectedSchemaId}
              placeholder="Choisir un schéma…"
              options={item.schemas.map((s) => ({ value: s.label, label: s.label }))}
            />
          </div>
        )}
      </div>
      <div style={{ flex: 1, minWidth: 0, width: "100%", maxWidth: isMobile ? "100%" : "600px" }}>
        {item.id === "huile" ? (
          item.driveFolderIdInstructionTechnique && (
            <DocumentSection label="Fiche de sécurité" driveFolderId={item.driveFolderIdInstructionTechnique} localFolderId={`${item.id}/instruction-technique`} produit={item.label} categorie={category.label} />
          )
        ) : (
          <>
            {item.driveFolderId && (
              <DocumentSection label="Notice d'utilisation" driveFolderId={item.driveFolderId} localFolderId={item.id} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdCodesDefaut && (
              <DocumentSection label="Codes défaut" driveFolderId={item.driveFolderIdCodesDefaut} localFolderId={`${item.id}/codes-defaut`} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdCommunication && (
              <DocumentSection label="Communication" driveFolderId={item.driveFolderIdCommunication} localFolderId={`${item.id}/communication`} forceDownload={true} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdUpdate && (
              <DocumentSection label="Update" driveFolderId={item.driveFolderIdUpdate} localFolderId={`${item.id}/update`} forceDownload={true} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdInstructionTechnique && (
              <DocumentSection label="Instruction technique" driveFolderId={item.driveFolderIdInstructionTechnique} localFolderId={`${item.id}/instruction-technique`} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdServiceInstruction && (
              <DocumentSection label="Service instruction" driveFolderId={item.driveFolderIdServiceInstruction} localFolderId={`${item.id}/service-instruction`} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdInstructionMontage && (
              <DocumentSection label="Instruction de montage" driveFolderId={item.driveFolderIdInstructionMontage} localFolderId={`${item.id}/instruction-montage`} produit={item.label} categorie={category.label} />
            )}
            {item.driveFolderIdDocuments && (
              <DocumentSection label="DIAGNOSTIC/ DEPANNAGE" driveFolderId={item.driveFolderIdDocuments} localFolderId={`${item.id}/documents`} produit={item.label} categorie={category.label} />
            )}
          </>
        )}
        <div style={{ display: "flex", justifyContent: "flex-start", gap: "14px", flexWrap: "wrap", marginTop: "28px" }}>
          {!simulateurInNoticeRow && <SimulateurButton item={item} onOpen={() => setShowSimulator(true)} />}
          <button onClick={() => setShowExpertForm(true)} style={{ display: "flex", alignItems: "center", gap: "8px", background: COLORS.gold, color: COLORS.navy, border: "none", borderRadius: "10px", padding: "14px 20px", fontSize: "13px", fontWeight: 700, letterSpacing: "0.02em", textTransform: "uppercase", cursor: "pointer" }}>
            Une question ? Contactez nos experts
            <ArrowRight size={16} />
          </button>
        </div>
      </div>
      {hasSchemas && (
        <>
          {/* Force le passage à la ligne pour afficher la zone de schéma sous le bouton "Une question ?..." plutôt qu'à côté */}
          <div style={{ flexBasis: "100%", width: 0 }} />
          <div style={{ flex: "3 1 256px", maxWidth: "64%", minWidth: isMobile ? "100%" : "256px", display: "flex", flexDirection: "column", alignItems: "center", gap: "10px", marginTop: "0px" }}>
            <div style={{ width: "100%", minHeight: "166px", background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "14px", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px" }}>
              {selectedSchema ? (
                <img src={selectedSchema.src} alt={selectedSchema.label} style={{ maxWidth: "100%", maxHeight: "70vh", objectFit: "contain" }} />
              ) : (
                <p style={{ fontSize: "13px", color: COLORS.textMuted, fontStyle: "italic", textAlign: "center", margin: 0 }}>Sélectionnez un schéma pour l'afficher ici.</p>
              )}
            </div>
            {selectedSchema && <p style={{ fontSize: "13px", fontWeight: 600, color: COLORS.navy, margin: 0 }}>{selectedSchema.label}</p>}
          </div>
        </>
      )}
      {showExpertForm && <ExpertRequestForm item={item} category={category} onClose={() => setShowExpertForm(false)} />}
      {showSimulator && <SimulatorViewer url={item.simulateurUrl} title={item.simulateurLabel || "Simulateur"} onClose={() => setShowSimulator(false)} />}
    </div>
  );
}

function SimulatorViewer({ url, title, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "#FFFFFF", display: "flex", flexDirection: "column", zIndex: 1000 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 20px", background: COLORS.navy, color: "#FFFFFF", flexShrink: 0 }}>
        <span style={{ fontSize: "14px", fontWeight: 600 }}>{title}</span>
        <button onClick={onClose} style={closeBtnStyle}><X size={16} />Fermer</button>
      </div>
      <iframe
        src={url}
        title={title}
        style={{ flex: 1, width: "100%", border: "none" }}
        allow="fullscreen"
      />
    </div>
  );
}

async function downloadFile(doc) {
  try {
    const res = await fetch(doc.url);
    if (!res.ok) throw new Error("Téléchargement impossible");
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = doc.name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(blobUrl);
  } catch (e) {
    window.open(doc.url, "_blank");
  }
}

function CustomSelect({ value, onChange, options, placeholder }) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef(null);

  useEffect(() => {
    if (!open) return;
    const handleOutside = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleOutside);
    document.addEventListener("touchstart", handleOutside);
    return () => {
      document.removeEventListener("mousedown", handleOutside);
      document.removeEventListener("touchstart", handleOutside);
    };
  }, [open]);

  const selected = options.find((o) => o.value === value);

  return (
    <div ref={containerRef} style={{ position: "relative", ...selectStyle, padding: 0, display: "flex", alignItems: "center" }}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        style={{
          width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "0 14px", background: "transparent", border: "none", color: selected ? COLORS.navy : COLORS.textMuted,
          fontSize: "13px", cursor: "pointer", textAlign: "left",
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{selected ? selected.label : placeholder}</span>
        <ChevronRight size={14} style={{ transform: open ? "rotate(90deg)" : "rotate(0deg)", transition: "transform 0.15s ease", flexShrink: 0, marginLeft: "8px" }} />
      </button>
      {open && (
        <div
          style={{
            position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, maxHeight: "260px", overflowY: "auto",
            background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "10px",
            boxShadow: "0 8px 24px rgba(11,31,58,0.15)", zIndex: 50,
          }}
        >
          {options.map((opt) => (
            <div
              key={opt.value}
              onClick={() => { onChange(opt.value); setOpen(false); }}
              style={{
                padding: "10px 14px", fontSize: "13px", cursor: "pointer", color: COLORS.navy,
                background: opt.value === value ? "#FBF1D9" : "transparent",
              }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "#FBF1D9"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = opt.value === value ? "#FBF1D9" : "transparent"; }}
            >
              {opt.label}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function FormationDocumentList({ driveFolderId, categorie }) {
  const [documents, setDocuments] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [viewingDoc, setViewingDoc] = useState(null);

  useEffect(() => {
    if (!driveFolderId) return;
    setLoading(true);
    setError(null);
    getDriveFiles(driveFolderId, "pdf")
      .then(setDocuments)
      .catch(() => setError("Impossible de charger les documents depuis Drive."))
      .finally(() => setLoading(false));
  }, [driveFolderId]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "10px", maxWidth: "480px" }}>
      {loading && <p style={{ fontSize: "13px", color: COLORS.textMuted }}>Chargement…</p>}
      {error && <p style={{ fontSize: "13px", color: "#C0392B" }}>{error}</p>}
      {!loading && !error && documents.length === 0 && <p style={{ fontSize: "13px", color: COLORS.textMuted, fontStyle: "italic" }}>Aucun document disponible.</p>}
      {!loading && !error && documents.map((docItem) => (
        <button
          key={docItem.id}
          onClick={() => setViewingDoc(docItem)}
          style={{ display: "flex", alignItems: "center", gap: "10px", background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "10px", padding: "12px 16px", fontSize: "13.5px", color: COLORS.navy, cursor: "pointer", textAlign: "left" }}
        >
          <FileText size={16} color={COLORS.gold} style={{ flexShrink: 0 }} />
          {docItem.name}
        </button>
      ))}
      {viewingDoc && <PdfViewer doc={viewingDoc} onClose={() => setViewingDoc(null)} categorie={categorie} />}
    </div>
  );
}

function GarantieCgvButton({ categorie }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [viewingDoc, setViewingDoc] = useState(null);

  const handleClick = async () => {
    setError(null);
    setLoading(true);
    try {
      const files = await getDriveFiles(GARANTIE_CGV_FOLDER_ID, "pdf");
      if (files.length === 0) throw new Error("Aucun document trouvé.");
      setViewingDoc(files[0]);
    } catch {
      setError("Impossible de charger le document pour l'instant.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <button
        onClick={handleClick}
        disabled={loading}
        style={{ display: "flex", alignItems: "center", gap: "10px", background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "10px", padding: "12px 16px", fontSize: "13.5px", color: COLORS.navy, cursor: loading ? "not-allowed" : "pointer", textAlign: "left", maxWidth: "480px" }}
      >
        <FileText size={16} color={COLORS.gold} style={{ flexShrink: 0 }} />
        {loading ? "Chargement…" : "Conditions générales de vente et de garantie"}
      </button>
      {error && <p style={{ fontSize: "13px", color: "#C0392B", marginTop: "8px" }}>{error}</p>}
      {viewingDoc && <PdfViewer doc={viewingDoc} onClose={() => setViewingDoc(null)} categorie={categorie} produit="Conditions générales de vente et de garantie" />}
    </>
  );
}

function DocumentSection({ label, driveFolderId, localFolderId, fileExtension = null, forceDownload = false, produit, categorie }) {
  const localDocuments = useMemo(() => getLocalDocuments(localFolderId), [localFolderId]);
  const [driveDocuments, setDriveDocuments] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [viewingDoc, setViewingDoc] = useState(null);
  const [selectedDocId, setSelectedDocId] = useState("");

  useEffect(() => {
    if (!driveFolderId) return;
    setLoading(true);
    setError(null);
    getDriveFiles(driveFolderId, fileExtension)
      .then(setDriveDocuments)
      .catch(() => setError("Impossible de charger les documents depuis Drive."))
      .finally(() => setLoading(false));
  }, [driveFolderId, fileExtension]);

  const documents = driveFolderId ? driveDocuments : localDocuments;
  
  const handleViewSelected = () => {
    const doc = documents.find((d) => d.id === selectedDocId);
    if (!doc) return;
    
    const isPdf = doc.name.toLowerCase().endsWith(".pdf");
    const isPreviewable = isPdf && !forceDownload;

    if (isPreviewable) {
      setViewingDoc(doc);
    } else {
      downloadFile(doc);
    }
  };

  return (
    <div style={{ marginTop: "22px" }}>
      <h3 style={{ fontSize: "12px", textTransform: "uppercase", letterSpacing: "0.06em", color: COLORS.textMuted, marginBottom: "10px" }}>{label}</h3>
      {loading && <p style={{ fontSize: "13px", color: COLORS.textMuted }}>Chargement…</p>}
      {error && <p style={{ fontSize: "13px", color: "#C0392B" }}>{error}</p>}
      {!loading && !error && documents.length === 0 && <p style={{ fontSize: "13px", color: COLORS.textMuted, fontStyle: "italic" }}>Aucun document disponible.</p>}
      {!loading && !error && documents.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: "8px" }}>
          <CustomSelect
            value={selectedDocId}
            onChange={setSelectedDocId}
            placeholder="Choisir un document…"
            options={documents.map((doc) => ({ value: doc.id, label: doc.name }))}
          />
          <button onClick={handleViewSelected} disabled={!selectedDocId} style={{ display: "flex", alignItems: "center", gap: "6px", background: selectedDocId ? COLORS.gold : "#EFEFEC", color: selectedDocId ? COLORS.navy : COLORS.textMuted, border: "none", borderRadius: "10px", padding: "0 18px", fontSize: "13px", fontWeight: 700, cursor: selectedDocId ? "pointer" : "default", whiteSpace: "nowrap" }}>
            {selectedDocId && documents.find(d => d.id === selectedDocId)?.name.toLowerCase().endsWith(".pdf") && !forceDownload ? <Eye size={16} /> : <FileText size={16} />}
            {selectedDocId && documents.find(d => d.id === selectedDocId)?.name.toLowerCase().endsWith(".pdf") && !forceDownload ? "Afficher" : "Télécharger"}
          </button>
        </div>
      )}
      {viewingDoc && <PdfViewer doc={viewingDoc} onClose={() => setViewingDoc(null)} produit={produit} categorie={categorie} />}
    </div>
  );
}

function ExpertRequestForm({ item, category, onClose, initialSujet = "Support Technique" }) {
  const [form, setForm] = useState({ nom: "", prenom: "", email: "", machine: "", numeroSerie: "", reference: "", sujet: initialSujet, message: "" });
  const hasCodesDefaut = Boolean(item?.driveFolderIdCodesDefaut);
  const [typeProbleme, setTypeProbleme] = useState("autres"); // "codes_defaut" | "autres"
  const [codeDefautChoisi, setCodeDefautChoisi] = useState("");
  const [motifAutres, setMotifAutres] = useState("");
  const [codesDefaut, setCodesDefaut] = useState([]);
  const [loadingCodesDefaut, setLoadingCodesDefaut] = useState(false);
  const [errorCodesDefaut, setErrorCodesDefaut] = useState(null);
  const [file, setFile] = useState(null);
  const messageEditorRef = useRef(null);
  const imageFilesRef = useRef(new Map()); // id -> File, source de vérité pour les images insérées dans le message
  const [imageCount, setImageCount] = useState(0); // uniquement pour l'affichage ("2 image(s) insérée(s)")
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState(null);
  const [sent, setSent] = useState(false);
  const update = (field) => (e) => setForm((f) => ({ ...f, [field]: e.target.value }));

  // Le nombre d'images est toujours recalculé à partir du contenu réel de
  // l'éditeur : si l'utilisateur supprime une image au clavier (sélection +
  // Suppr), le compteur et la liste envoyée à la fin restent synchronisés.
  const syncImageCount = () => {
    setImageCount(messageEditorRef.current?.querySelectorAll("img[data-image-id]").length || 0);
  };

  const insertImageIntoEditor = (id, previewUrl) => {
    const editor = messageEditorRef.current;
    if (!editor) return;
    editor.focus();
    const html = `<img src="${previewUrl}" data-image-id="${id}" alt="" style="max-width:140px;max-height:140px;border-radius:8px;vertical-align:middle;margin:3px;object-fit:cover;cursor:pointer;" title="Double-cliquez pour retirer l'image" />`;
    document.execCommand("insertHTML", false, html);
    syncImageCount();
  };

  const addImages = (fileList) => {
    const newFiles = Array.from(fileList || []).filter((f) => f.type.startsWith("image/"));
    newFiles.forEach((f) => {
      const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
      const previewUrl = URL.createObjectURL(f);
      imageFilesRef.current.set(id, f);
      insertImageIntoEditor(id, previewUrl);
    });
  };

  // Double-clic sur une image insérée dans le message pour la retirer.
  const handleEditorDoubleClick = (e) => {
    if (e.target.tagName === "IMG") {
      const id = e.target.getAttribute("data-image-id");
      URL.revokeObjectURL(e.target.src);
      e.target.remove();
      if (id) imageFilesRef.current.delete(id);
      syncImageCount();
    }
  };

  // Permet de coller directement une capture d'écran (Ctrl+V) dans la zone de message.
  // Tout collage non-image est forcé en texte brut : le contenu de l'éditeur est
  // enregistré tel quel (HTML) pour préserver la position des images, il ne doit
  // donc jamais contenir de balises arbitraires collées depuis une autre page.
  const handleMessagePaste = (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    const pastedFiles = [];
    for (const item of items) {
      if (item.kind === "file" && item.type.startsWith("image/")) {
        const f = item.getAsFile();
        if (f) pastedFiles.push(f);
      }
    }
    e.preventDefault();
    if (pastedFiles.length) {
      addImages(pastedFiles);
    } else {
      const text = e.clipboardData.getData("text/plain");
      if (text) document.execCommand("insertText", false, text);
    }
  };

  // Suggère Nom / Prénom / E-mail à partir du compte connecté : l'adresse
  // e-mail professionnelle suit le format prenom.nom@domaine, on en déduit
  // le nom et le prénom sans écraser une valeur déjà saisie par l'utilisateur.
  useEffect(() => {
    const currentUser = auth.currentUser;
    const email = currentUser?.email || "";
    const localPart = email.split("@")[0] || "";
    const [prenomPart, ...resteParts] = localPart.split(/[._-]+/).filter(Boolean);
    const nomPart = resteParts.join(" ");
    const capitalize = (s) => s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : "";
    setForm((f) => ({
      ...f,
      prenom: f.prenom || capitalize(prenomPart),
      nom: f.nom || capitalize(nomPart),
      email: f.email || email,
    }));
  }, []);

  useEffect(() => {
    if (typeProbleme !== "codes_defaut" || !hasCodesDefaut) return;
    setLoadingCodesDefaut(true);
    setErrorCodesDefaut(null);
    getDriveFiles(item.driveFolderIdCodesDefaut, "pdf")
      .then(setCodesDefaut)
      .catch(() => setErrorCodesDefaut("Impossible de charger la liste des codes défaut."))
      .finally(() => setLoadingCodesDefaut(false));
  }, [typeProbleme, item, hasCodesDefaut]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError(null);
    if (typeProbleme === "codes_defaut" && !codeDefautChoisi) {
      setError("Veuillez sélectionner un code défaut.");
      return;
    }
    if (typeProbleme === "autres" && !motifAutres.trim()) {
      setError("Veuillez indiquer le motif de votre demande.");
      return;
    }
    setSubmitting(true);
    try {
      const motifFinal = typeProbleme === "codes_defaut"
        ? codeDefautChoisi.replace(/\.pdf$/i, "")
        : motifAutres.trim();
      const messageTexte = (messageEditorRef.current?.innerText || "").trim();
      const imageNodes = messageEditorRef.current
        ? Array.from(messageEditorRef.current.querySelectorAll("img[data-image-id]"))
        : [];
      const formAEnvoyer = { ...form, message: messageTexte, motif: motifFinal };
      const destinataire = SERVICE_EMAILS[form.sujet] || null;
      let pieceJointeUrl = null;
      if (file) {
        pieceJointeUrl = await uploadFileToDrive(file);
      }
      let imageUploadFailed = false;
      const imagesMessage = imageNodes.length
        ? await Promise.all(imageNodes.map(async (node) => {
            const id = node.getAttribute("data-image-id");
            const imgFile = imageFilesRef.current.get(id);
            if (!imgFile) { node.remove(); return null; }
            const originalSrc = node.src;
            try {
              const url = await uploadFileToDrive(imgFile);
              // Remplace l'aperçu local par l'URL Drive définitive, à la même position dans le message.
              node.src = url;
              node.removeAttribute("data-image-id");
              node.removeAttribute("title");
              URL.revokeObjectURL(originalSrc);
              return { nom: imgFile.name, url };
            } catch (err) {
              // Une image en échec ne doit jamais rester avec son URL blob:
              // locale (invalide hors de cette session) ni bloquer l'envoi
              // des autres images ou du reste de la demande.
              console.error("Erreur upload image message:", err);
              imageUploadFailed = true;
              node.remove();
              URL.revokeObjectURL(originalSrc);
              return null;
            }
          }))
        : [];
      const imagesMessageFiltrees = imagesMessage.filter(Boolean);
      const messageHtml = messageEditorRef.current ? messageEditorRef.current.innerHTML : "";
      const docRef = await addDoc(collection(db, "expertRequests"), { ...formAEnvoyer, messageHtml, destinataire, pieceJointeNom: file?.name || null, pieceJointeUrl, imagesMessage: imagesMessageFiltrees, produit: item?.label || category?.label || null, categorie: category?.label || null, requestedBy: auth.currentUser?.email || null, createdAt: serverTimestamp(), status: "En attente", archived: false });
      await fetch(APPS_SCRIPT_URL, {
        method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ type: "nouvelle_demande", requestId: docRef.id, ...formAEnvoyer, destinataire, produit: item?.label || category?.label || "", categorie: category?.label || "", pieceJointe: file?.name || "Aucune", imagesMessage: imagesMessageFiltrees.map((i) => i.nom).join(", ") || "Aucune" }),
      });
      setSent(true);
      if (imageUploadFailed) {
        setError("Votre demande a bien été envoyée, mais une ou plusieurs images n'ont pas pu être jointes.");
      }
    } catch (err) {
      setError(err?.message && (file || imageFilesRef.current.size > 0) ? `Pièce jointe : ${err.message}` : "Impossible d'envoyer votre demande pour l'instant. Réessayez.");
    } finally {
      setSubmitting(false);
    }
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: "20px" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#FFFFFF", borderRadius: "16px", maxWidth: "680px", width: "100%", maxHeight: "90vh", overflow: "auto", padding: "28px", position: "relative" }}>
        <button onClick={onClose} aria-label="Fermer" style={{ position: "absolute", top: "16px", right: "16px", background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={20} /></button>
        {sent ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <h2 style={{ fontSize: "18px", fontWeight: 700, marginBottom: "10px" }}>Demande envoyée</h2>
            <p style={{ color: COLORS.textMuted, fontSize: "14px", marginBottom: "20px" }}>Un expert vous répondra dans les meilleurs délais.</p>
            <button onClick={onClose} style={primaryBtnStyle}>Fermer</button>
          </div>
        ) : (
          <>
            <h2 style={{ fontSize: "20px", fontWeight: 800, marginBottom: "18px" }}>Nouvelle Demande d'Expertise</h2>
            <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
              <input required placeholder="Nom" value={form.nom} onChange={update("nom")} style={formInputStyle} />
              <input required placeholder="Prénom" value={form.prenom} onChange={update("prenom")} style={formInputStyle} />
              <input required type="email" placeholder="E-mail" value={form.email} onChange={update("email")} style={formInputStyle} />
              <FieldWithTooltip placeholder="Machine" value={form.machine} onChange={update("machine")} tooltip="Le modèle exact de la machine concernée (ex. CSDX SFC)." tooltipImage={imgTooltipMachine} />
              <FieldWithTooltip placeholder="N° Série" value={form.numeroSerie} onChange={update("numeroSerie")} tooltip="Le numéro de série figure sur la plaque signalétique de la machine." tooltipImage={imgTooltipNumeroSerie} />
              <FieldWithTooltip placeholder="Référence" value={form.reference} onChange={update("reference")} tooltip="Une référence interne de commande ou de dossier, si vous en avez une." tooltipImage={imgTooltipReference} />
              <select required value={form.sujet} onChange={update("sujet")} style={formInputStyle}>
                <option>Support Technique</option><option>Garantie</option><option>Pièces détachées</option><option>Formation</option>
              </select>
              {hasCodesDefaut && (
                <div>
                  <label style={{ fontSize: "12px", fontWeight: 700, color: COLORS.navy, marginBottom: "6px", display: "block" }}>Motif</label>
                  <select
                    required
                    value={typeProbleme}
                    onChange={(e) => { setTypeProbleme(e.target.value); setCodeDefautChoisi(""); }}
                    style={formInputStyle}
                  >
                    <option value="autres">Autres</option>
                    <option value="codes_defaut">Codes défauts</option>
                  </select>
                </div>
              )}
              {typeProbleme === "codes_defaut" && hasCodesDefaut && (
                <div>
                  {loadingCodesDefaut && <p style={{ fontSize: "13px", color: COLORS.textMuted }}>Chargement des codes défaut…</p>}
                  {errorCodesDefaut && <p style={{ fontSize: "13px", color: "#C0392B" }}>{errorCodesDefaut}</p>}
                  {!loadingCodesDefaut && !errorCodesDefaut && codesDefaut.length > 0 && (
                    <select required value={codeDefautChoisi} onChange={(e) => setCodeDefautChoisi(e.target.value)} style={formInputStyle}>
                      <option value="" disabled>Sélectionner un code défaut…</option>
                      {codesDefaut.map((c) => (
                        <option key={c.id} value={c.name}>{c.name.replace(/\.pdf$/i, "")}</option>
                      ))}
                    </select>
                  )}
                  {!loadingCodesDefaut && !errorCodesDefaut && codesDefaut.length === 0 && (
                    <p style={{ fontSize: "13px", color: COLORS.textMuted, fontStyle: "italic" }}>Aucun code défaut disponible pour ce produit.</p>
                  )}
                </div>
              )}
              {typeProbleme === "autres" && (
                <div>
                  {!hasCodesDefaut && <label style={{ fontSize: "12px", fontWeight: 700, color: COLORS.navy, marginBottom: "6px", display: "block" }}>Motif</label>}
                  <input
                    required
                    type="text"
                    placeholder="Motif de votre demande"
                    value={motifAutres}
                    onChange={(e) => setMotifAutres(e.target.value)}
                    style={formInputStyle}
                  />
                </div>
              )}
              <div>
                <style>{`
                  .expert-message-editor:empty:before {
                    content: attr(data-placeholder);
                    color: ${COLORS.textMuted};
                    pointer-events: none;
                  }
                  .expert-message-editor img:hover {
                    outline: 2px solid ${COLORS.gold};
                  }
                `}</style>
                <div
                  ref={messageEditorRef}
                  className="expert-message-editor"
                  contentEditable
                  suppressContentEditableWarning
                  onPaste={handleMessagePaste}
                  onInput={syncImageCount}
                  onDoubleClick={handleEditorDoubleClick}
                  data-placeholder="Votre message (informations complémentaires, optionnel)..."
                  style={{ ...formInputStyle, minHeight: "180px", maxHeight: "420px", overflowY: "auto", fontFamily: "inherit", lineHeight: 1.5, cursor: "text" }}
                />
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginTop: "6px", flexWrap: "wrap" }}>
                  <span style={{ fontSize: "11.5px", color: COLORS.textMuted, fontStyle: "italic" }}>
                    Collez une capture d'écran (Ctrl+V) directement dans le message ci-dessus
                  </span>
                </div>
                {imageCount > 0 && (
                  <p style={{ fontSize: "11.5px", color: COLORS.textMuted, marginTop: "4px" }}>
                    {imageCount} image{imageCount > 1 ? "s" : ""} insérée{imageCount > 1 ? "s" : ""} — double-cliquez sur une image pour la retirer.
                  </p>
                )}
              </div>
              <div>
                <label style={{ fontSize: "12px", color: COLORS.textMuted, display: "flex", alignItems: "center", gap: "6px", marginBottom: "6px" }}><Paperclip size={14} />Pièce jointe (optionnel)</label>
                <input type="file" onChange={(e) => setFile(e.target.files?.[0] || null)} style={{ fontSize: "13px" }} />
              </div>
              {error && <p style={{ color: "#C0392B", fontSize: "13px" }}>{error}</p>}
              <button type="submit" disabled={submitting} style={{ ...primaryBtnStyle, opacity: submitting ? 0.7 : 1, marginTop: "8px" }}>{submitting ? "Envoi…" : "Envoyer"}</button>
            </form>
          </>
        )}
      </div>
    </div>
  );
}

function MachineRegistrationForm({ category, onClose }) {
  const [form, setForm] = useState({ nom: "", prenom: "", email: "", clientFinal: "", adresse: "", machine: "", numeroSerie: "", reference: "", garantie: "", message: "" });
  const [ficheMiseEnRoute, setFicheMiseEnRoute] = useState("");
  const [photosInstallation, setPhotosInstallation] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState(null);
  const [sent, setSent] = useState(false);
  const update = (field) => (e) => setForm((f) => ({ ...f, [field]: e.target.value }));
  const handleSubmit = async (e) => {
    e.preventDefault();
    setError(null);
    setSubmitting(true);
    try {
      const destinataire = SERVICE_EMAILS[category?.label] || SERVICE_EMAILS["Garantie"];
      await addDoc(collection(db, "machineRegistrations"), { ...form, destinataire, ficheMiseEnRoute: ficheMiseEnRoute || null, photosInstallation: photosInstallation || null, categorie: category?.label || null, requestedBy: auth.currentUser?.email || null, createdAt: serverTimestamp() });
      await fetch(APPS_SCRIPT_URL, {
        method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ ...form, destinataire, type: "Enregistrement Machine", categorie: category?.label || "", ficheMiseEnRoute: ficheMiseEnRoute || "Aucune", photosInstallation: photosInstallation || "Aucune" }),
      });
      setSent(true);
    } catch {
      setError("Impossible d'envoyer votre demande pour l'instant. Réessayez.");
    } finally {
      setSubmitting(false);
    }
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: "20px" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#FFFFFF", borderRadius: "16px", maxWidth: "480px", width: "100%", maxHeight: "90vh", overflow: "auto", padding: "28px", position: "relative" }}>
        <button onClick={onClose} aria-label="Fermer" style={{ position: "absolute", top: "16px", right: "16px", background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={20} /></button>
        {sent ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <h2 style={{ fontSize: "18px", fontWeight: 700, marginBottom: "10px" }}>Machine enregistrée</h2>
            <p style={{ color: COLORS.textMuted, fontSize: "14px", marginBottom: "20px" }}>Votre enregistrement a bien été pris en compte.</p>
            <button onClick={onClose} style={primaryBtnStyle}>Fermer</button>
          </div>
        ) : (
          <>
            <h2 style={{ fontSize: "20px", fontWeight: 800, marginBottom: "18px" }}>Enregistrement Machine</h2>
            <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
              <input required placeholder="Nom" value={form.nom} onChange={update("nom")} style={formInputStyle} />
              <input required placeholder="Prénom" value={form.prenom} onChange={update("prenom")} style={formInputStyle} />
              <input required type="email" placeholder="E-mail" value={form.email} onChange={update("email")} style={formInputStyle} />
              <input required placeholder="Nom du client final" value={form.clientFinal} onChange={update("clientFinal")} style={formInputStyle} />
              <input required placeholder="Adresse postale" value={form.adresse} onChange={update("adresse")} style={formInputStyle} />
              <FieldWithTooltip placeholder="Machine" value={form.machine} onChange={update("machine")} tooltip="Le modèle exact de la machine concernée (ex. CSDX SFC)." tooltipImage={imgTooltipMachine} />
              <FieldWithTooltip placeholder="N° Série" value={form.numeroSerie} onChange={update("numeroSerie")} tooltip="Le numéro de série figure sur la plaque signalétique de la machine." tooltipImage={imgTooltipNumeroSerie} />
              <FieldWithTooltip placeholder="Référence" value={form.reference} onChange={update("reference")} tooltip="Une référence interne de commande ou de dossier, si vous en avez une." tooltipImage={imgTooltipReference} />
              <select required value={form.garantie} onChange={update("garantie")} style={formInputStyle}>
                <option value="" disabled>Garantie</option>
                <option>Garantie standard</option>
                <option>Garantie étendue</option>
              </select>
              <textarea placeholder="Votre message..." value={form.message} onChange={update("message")} rows={4} style={{ ...formInputStyle, resize: "vertical", fontFamily: "inherit" }} />
              <div>
                <label style={{ fontSize: "12px", color: COLORS.textMuted, display: "flex", alignItems: "center", gap: "6px", marginBottom: "6px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>Fiche mise en route</label>
                <input type="file" onChange={(e) => setFicheMiseEnRoute(e.target.files?.[0]?.name || "")} style={{ fontSize: "13px" }} />
              </div>
              <div>
                <label style={{ fontSize: "12px", color: COLORS.textMuted, display: "flex", alignItems: "center", gap: "6px", marginBottom: "6px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>Photos d'installation</label>
                <input type="file" onChange={(e) => setPhotosInstallation(e.target.files?.[0]?.name || "")} style={{ fontSize: "13px" }} />
              </div>
              {error && <p style={{ color: "#C0392B", fontSize: "13px" }}>{error}</p>}
              <button type="submit" disabled={submitting} style={{ ...primaryBtnStyle, opacity: submitting ? 0.7 : 1, marginTop: "8px" }}>{submitting ? "Envoi…" : "Envoyer"}</button>
            </form>
          </>
        )}
      </div>
    </div>
  );
}

function FieldWithTooltip({ tooltip, tooltipImage, ...inputProps }) {
  const [showTooltip, setShowTooltip] = useState(false);
  const [coords, setCoords] = useState({ top: 0, left: 0 });
  const iconRef = useRef(null);

  const handleEnter = () => {
    if (iconRef.current) {
      const rect = iconRef.current.getBoundingClientRect();
      setCoords({ top: rect.bottom + 8, left: rect.right });
    }
    setShowTooltip(true);
  };

  return (
    <div style={{ position: "relative" }}>
      <input required {...inputProps} style={{ ...formInputStyle, paddingRight: "36px" }} />
      <span
        ref={iconRef}
        onMouseEnter={handleEnter}
        onMouseLeave={() => setShowTooltip(false)}
        style={{ position: "absolute", right: "10px", top: "50%", transform: "translateY(-50%)", cursor: "help", display: "flex" }}
      >
        <img src={iconInfo} alt="Info" style={{ width: 16, height: 16 }} />
      </span>
      {showTooltip && tooltipImage && createPortal(
        <div style={{ position: "fixed", top: coords.top, left: coords.left, transform: "translateX(-100%)", zIndex: 9999, background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "12px", padding: "12px", boxShadow: "0 16px 40px rgba(11,31,58,0.35)", pointerEvents: "none" }}>
          <img src={tooltipImage} alt={tooltip || ""} style={{ width: 380, maxWidth: "80vw", height: "auto", display: "block" }} />
        </div>,
        document.body
      )}
    </div>
  );
}

const formInputStyle = { width: "100%", padding: "12px 14px", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}`, fontSize: "14px", outline: "none", boxSizing: "border-box" };
const primaryBtnStyle = { background: COLORS.gold, color: COLORS.navy, border: "none", borderRadius: "10px", padding: "14px 20px", fontSize: "13px", fontWeight: 700, cursor: "pointer" };

function logDocumentConsultation(doc, produit, categorie) {
  try {
    addDoc(collection(db, "documentConsultations"), {
      fileName: doc?.name || "Document inconnu",
      viewedBy: auth.currentUser?.email || null,
      produit: produit || null,
      categorie: categorie || null,
      viewedAt: serverTimestamp(),
    }).catch(() => {});
    fetch(APPS_SCRIPT_URL, {
      method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        type: "consultation_document",
        fileName: doc?.name || "Document inconnu",
        utilisateur: auth.currentUser?.email || "",
        produit: produit || "",
        categorie: categorie || "",
      }),
    }).catch(() => {});
  } catch {
    // La journalisation de la consultation ne doit jamais empêcher l'affichage du PDF.
  }
}

function PdfViewer({ doc, onClose, produit, categorie }) {
  const isMobile = useIsMobile();
  const [numPages, setNumPages] = useState(null);
  const [pageNumber, setPageNumber] = useState(1);
  const [showOutline, setShowOutline] = useState(() => typeof window !== "undefined" && window.innerWidth > 640);
  const [hasOutline, setHasOutline] = useState(true);
  const [rotation, setRotation] = useState(0);
  const [zoom, setZoom] = useState(1);
  const ZOOM_MIN = 0.5;
  const ZOOM_MAX = 2.5;
  const ZOOM_STEP = 0.25;
  const consultationLoggedRef = useRef(false);
  useEffect(() => { const block = (e) => e.preventDefault(); document.addEventListener("contextmenu", block); return () => document.removeEventListener("contextmenu", block); }, []);
  useEffect(() => {
    if (consultationLoggedRef.current) return;
    consultationLoggedRef.current = true;
    logDocumentConsultation(doc, produit, categorie);
  }, []);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.85)", display: "flex", flexDirection: "column", zIndex: 1000 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "8px", padding: isMobile ? "10px 12px" : "14px 20px", background: COLORS.navy, color: "#FFFFFF" }}>
        <span style={{ fontSize: "13px", fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: isMobile ? "100%" : "40%" }}>{doc.name}</span>
        <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: isMobile ? "8px" : "14px" }}>
          <button onClick={() => setShowOutline((s) => !s)} style={{ ...navBtnStyle, background: showOutline ? COLORS.gold : "rgba(255,255,255,0.15)" }} title="Signets"><PanelLeft size={16} color={showOutline ? COLORS.navy : "#FFFFFF"} /></button>
          {numPages && (
            <div style={{ display: "flex", alignItems: "center", gap: "8px", fontSize: "13px" }}>
              <button onClick={() => setPageNumber((p) => Math.max(1, p - 1))} disabled={pageNumber <= 1} style={navBtnStyle}><ChevronLeft size={16} /></button>
              <span>{isMobile ? `${pageNumber}/${numPages}` : `Page ${pageNumber} / ${numPages}`}</span>
              <button onClick={() => setPageNumber((p) => Math.min(numPages, p + 1))} disabled={pageNumber >= numPages} style={navBtnStyle}><ChevronRight size={16} /></button>
            </div>
          )}
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <button onClick={() => setZoom((z) => Math.max(ZOOM_MIN, +(z - ZOOM_STEP).toFixed(2)))} disabled={zoom <= ZOOM_MIN} style={navBtnStyle} title="Dézoomer"><ZoomOut size={16} /></button>
            <span style={{ fontSize: "13px", minWidth: "40px", textAlign: "center" }}>{Math.round(zoom * 100)}%</span>
            <button onClick={() => setZoom((z) => Math.min(ZOOM_MAX, +(z + ZOOM_STEP).toFixed(2)))} disabled={zoom >= ZOOM_MAX} style={navBtnStyle} title="Zoomer"><ZoomIn size={16} /></button>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <button onClick={() => setRotation((r) => (r - 90 + 360) % 360)} style={navBtnStyle} title="Pivoter à gauche"><RotateCcw size={16} /></button>
            <button onClick={() => setRotation((r) => (r + 90) % 360)} style={navBtnStyle} title="Pivoter à droite"><RotateCw size={16} /></button>
          </div>
          <button onClick={onClose} style={closeBtnStyle}><X size={16} />{!isMobile && "Fermer"}</button>
        </div>
      </div>
      <div style={{ flex: 1, display: "flex", overflow: "hidden", position: "relative" }}>
        <Document file={doc.url} onLoadSuccess={(pdf) => setNumPages(pdf.numPages)} loading={<p style={{ color: "#FFFFFF", padding: "24px" }}>Chargement du document…</p>} error={<p style={{ color: "#FFFFFF", padding: "24px" }}>Impossible d'afficher ce PDF.</p>}>
          <div style={{ display: "flex", flex: 1, height: "100%", position: "relative" }}>
            {showOutline && (
              <div
                style={
                  isMobile
                    ? { position: "absolute", inset: 0, zIndex: 5, width: "85vw", maxWidth: "300px", overflow: "auto", background: "#FFFFFF", borderRight: `1px solid ${COLORS.cardBorder}`, padding: "16px", boxShadow: "4px 0 16px rgba(0,0,0,0.25)" }
                    : { width: "260px", flexShrink: 0, overflow: "auto", background: "#FFFFFF", borderRight: `1px solid ${COLORS.cardBorder}`, padding: "16px" }
                }
              >
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "10px" }}>
                  <h4 style={{ fontSize: "11px", textTransform: "uppercase", letterSpacing: "0.06em", color: COLORS.textMuted, margin: 0 }}>Signets</h4>
                  {isMobile && <button onClick={() => setShowOutline(false)} aria-label="Fermer les signets" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={16} /></button>}
                </div>
                <div style={{ fontSize: "12px", lineHeight: 1.5 }}><Outline onLoadSuccess={(outline) => setHasOutline(Boolean(outline && outline.length))} onItemClick={({ pageNumber: pn }) => { setPageNumber(pn); if (isMobile) setShowOutline(false); }} /></div>
                {!hasOutline && <p style={{ fontSize: "12px", color: COLORS.textMuted }}>Ce PDF ne contient pas de signets.</p>}
              </div>
            )}
            <div onContextMenu={(e) => e.preventDefault()} style={{ flex: 1, overflow: "auto", display: "flex", justifyContent: "center", padding: isMobile ? "12px" : "24px", background: "#5A5A5A" }}>
              <Page pageNumber={pageNumber} rotate={rotation} renderAnnotationLayer={false} renderTextLayer={false} width={Math.min(1400, window.innerWidth - (!isMobile && showOutline ? 300 : 24)) * zoom} />
            </div>
          </div>
        </Document>
      </div>
    </div>
  );
}

function Card({ icon: Icon, image, label, description, iconColor, onClick, imageSize }) {
  const imgSize = imageSize || (description ? 120 : 140);
  return (
    <button onClick={onClick} style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", textAlign: "center", gap: "14px", padding: description ? "22px 20px" : "20px 16px", borderRadius: "14px", border: `1px solid ${COLORS.cardBorder}`, background: "#FFFFFF", color: COLORS.navy, cursor: "pointer", boxShadow: "0 1px 3px rgba(11,31,58,0.04)", transition: "transform 0.15s ease, box-shadow 0.15s ease" }} onMouseEnter={(e) => { e.currentTarget.style.transform = "translateY(-3px)"; e.currentTarget.style.boxShadow = "0 8px 20px rgba(11,31,58,0.08)"; }} onMouseLeave={(e) => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "0 1px 3px rgba(11,31,58,0.04)"; }}>
      {image ? <img src={image} alt="" style={{ width: imgSize, height: imgSize, objectFit: "contain" }} /> : <Icon size={description ? 26 : 28} color={iconColor || COLORS.gold} strokeWidth={1.6} />}
      <div>
        <div style={{ fontSize: "14px", fontWeight: 700, letterSpacing: "0.01em", textTransform: description ? "uppercase" : "none" }}>{label}</div>
        {description && <div style={{ fontSize: "12.5px", color: COLORS.textMuted, marginTop: "6px" }}>{description}</div>}
      </div>
    </button>
  );
}

function PdfNotificationsModal({ notifications, currentUserEmail, onClose, onForceScan, scanning }) {
  const [checked, setChecked] = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState(null);
  const [scanMessage, setScanMessage] = useState(null);

  const toggle = (id) => setChecked((c) => ({ ...c, [id]: !c[id] }));
  const selectedIds = Object.keys(checked).filter((id) => checked[id]);
  const allSelected = notifications.length > 0 && selectedIds.length === notifications.length;

  const toggleSelectAll = () => {
    if (allSelected) {
      setChecked({});
    } else {
      setChecked(Object.fromEntries(notifications.map((n) => [n.id, true])));
    }
  };

  const handleForceScan = async () => {
    setScanMessage(null);
    const result = await onForceScan();
    if (!result?.ok) {
      setScanMessage("La vérification a échoué. Ouvrez la console du navigateur (F12) pour voir le détail de l'erreur.");
    } else if (result.created > 0) {
      setScanMessage(`${result.created} nouveau(x) document(s) trouvé(s).`);
    } else {
      setScanMessage("Aucun nouveau document trouvé.");
    }
  };

  const validateIds = async (ids) => {
    if (ids.length === 0) return;
    setSubmitting(true);
    setError(null);
    try {
      await Promise.all(ids.map((id) => updateDoc(doc(db, "pdfNotifications", id), { validatedBy: arrayUnion(currentUserEmail) })));
      setChecked({});
    } catch {
      setError("Impossible d'enregistrer la validation pour l'instant. Réessayez.");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: "20px" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#FFFFFF", borderRadius: "16px", maxWidth: "560px", width: "100%", maxHeight: "85vh", overflow: "auto", padding: "28px", position: "relative" }}>
        <button onClick={onClose} aria-label="Fermer" style={{ position: "absolute", top: "16px", right: "16px", background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={20} /></button>
        <h2 style={{ fontSize: "20px", fontWeight: 800, marginBottom: "6px", display: "flex", alignItems: "center", gap: "8px" }}>
          <Bell size={20} color={COLORS.gold} />Nouveaux documents disponibles
        </h2>
        <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: notifications.length === 0 ? 0 : "4px" }}>
          <button onClick={handleForceScan} disabled={scanning} style={{ display: "flex", alignItems: "center", gap: "6px", background: "none", border: "none", color: COLORS.goldDark, fontSize: "12.5px", fontWeight: 700, cursor: "pointer", padding: 0, opacity: scanning ? 0.6 : 1 }}>
            <RefreshCw size={13} />{scanning ? "Vérification…" : "Vérifier maintenant"}
          </button>
        </div>
        {scanMessage && <p style={{ fontSize: "12px", color: COLORS.textMuted, marginBottom: "10px" }}>{scanMessage}</p>}
        {notifications.length === 0 ? (
          <>
            <p style={{ fontSize: "14px", color: COLORS.textMuted, marginTop: "16px" }}>Vous êtes à jour, aucune notification en attente.</p>
            <button onClick={onClose} style={{ ...primaryBtnStyle, marginTop: "18px" }}>Fermer</button>
          </>
        ) : (
          <>
            <p style={{ fontSize: "13px", color: COLORS.textMuted, marginBottom: "14px" }}>
              De nouveaux documents PDF ont été ajoutés. Cochez ceux dont vous avez pris connaissance puis validez : tant qu'un document n'est pas validé, il reste affiché ici.
            </p>
            {error && <p style={{ color: "#C0392B", fontSize: "13px", marginBottom: "10px" }}>{error}</p>}
            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "16px" }}>
              <button onClick={() => validateIds(selectedIds)} disabled={submitting || selectedIds.length === 0} style={{ ...primaryBtnStyle, opacity: submitting || selectedIds.length === 0 ? 0.6 : 1 }}>
                {submitting ? "Validation…" : "Valider la sélection"}
              </button>
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: "8px", padding: "10px 14px", marginBottom: "10px", borderRadius: "10px", background: "#FAFAF8", border: `1px solid ${COLORS.cardBorder}`, cursor: "pointer", fontSize: "13px", fontWeight: 700, color: COLORS.navy }}>
              <input type="checkbox" checked={allSelected} onChange={toggleSelectAll} />
              Tout sélectionner ({notifications.length})
            </label>
            <div style={{ display: "flex", flexDirection: "column", gap: "10px", marginBottom: "6px" }}>
              {notifications.map((n) => (
                <label key={n.id} style={{ display: "flex", alignItems: "flex-start", gap: "10px", padding: "12px 14px", borderRadius: "10px", border: `1px solid ${COLORS.cardBorder}`, cursor: "pointer" }}>
                  <input type="checkbox" checked={!!checked[n.id]} onChange={() => toggle(n.id)} style={{ marginTop: "3px" }} />
                  <span>
                    <span style={{ display: "flex", alignItems: "center", gap: "6px", fontWeight: 700, fontSize: "13.5px", color: COLORS.navy }}>
                      <FileText size={14} color={COLORS.gold} />{n.fileName}
                    </span>
                    {n.context && <span style={{ display: "block", fontSize: "12px", color: COLORS.textMuted, marginTop: "2px" }}>{n.context}</span>}
                  </span>
                </label>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function formatRequestDate(ts) {
  const d = ts?.toDate?.();
  if (!d) return "";
  return d.toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

function ExpertRequestCard({ request, isAdmin, onArchive, archiving }) {
  const isProcessed = !!request.archived;
  const [reponseFile, setReponseFile] = useState(null);
  const reponseEditorRef = useRef(null);
  const reponseImageFilesRef = useRef(new Map()); // id -> File, images insérées dans la réponse
  const [reponseImageCount, setReponseImageCount] = useState(0);
  const [reponseHasText, setReponseHasText] = useState(false);
  const canSend = reponseHasText || reponseImageCount > 0;

  const syncReponseState = () => {
    const editor = reponseEditorRef.current;
    if (!editor) return;
    setReponseImageCount(editor.querySelectorAll("img[data-image-id]").length);
    setReponseHasText(editor.innerText.trim().length > 0);
  };

  const insertReponseImage = (id, previewUrl) => {
    const editor = reponseEditorRef.current;
    if (!editor) return;
    editor.focus();
    const html = `<img src="${previewUrl}" data-image-id="${id}" alt="" style="max-width:140px;max-height:140px;border-radius:8px;vertical-align:middle;margin:3px;object-fit:cover;cursor:pointer;" title="Double-cliquez pour retirer l'image" />`;
    document.execCommand("insertHTML", false, html);
    syncReponseState();
  };

  const addReponseImages = (fileList) => {
    const newFiles = Array.from(fileList || []).filter((f) => f.type.startsWith("image/"));
    newFiles.forEach((f) => {
      const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
      const previewUrl = URL.createObjectURL(f);
      reponseImageFilesRef.current.set(id, f);
      insertReponseImage(id, previewUrl);
    });
  };

  // Double-clic sur une image insérée dans la réponse pour la retirer.
  const handleReponseEditorDoubleClick = (e) => {
    if (e.target.tagName === "IMG") {
      const id = e.target.getAttribute("data-image-id");
      URL.revokeObjectURL(e.target.src);
      e.target.remove();
      if (id) reponseImageFilesRef.current.delete(id);
      syncReponseState();
    }
  };

  // Permet de coller directement une capture d'écran (Ctrl+V) dans la réponse.
  // Tout collage non-image est forcé en texte brut (le HTML de la réponse est
  // enregistré tel quel pour préserver la position des images).
  const handleReponsePaste = (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    const pastedFiles = [];
    for (const item of items) {
      if (item.kind === "file" && item.type.startsWith("image/")) {
        const f = item.getAsFile();
        if (f) pastedFiles.push(f);
      }
    }
    e.preventDefault();
    if (pastedFiles.length) {
      addReponseImages(pastedFiles);
    } else {
      const text = e.clipboardData.getData("text/plain");
      if (text) document.execCommand("insertText", false, text);
    }
  };

  const handleSend = () => {
    const texte = (reponseEditorRef.current?.innerText || "").trim();
    const html = reponseEditorRef.current ? reponseEditorRef.current.innerHTML : "";
    const imageNodes = reponseEditorRef.current
      ? Array.from(reponseEditorRef.current.querySelectorAll("img[data-image-id]"))
      : [];
    const images = imageNodes
      .map((node) => ({ id: node.getAttribute("data-image-id"), file: reponseImageFilesRef.current.get(node.getAttribute("data-image-id")) }))
      .filter((img) => img.file);
    onArchive(request, texte, reponseFile, images, html);
  };

  return (
    <div style={{ padding: "14px 16px", borderRadius: "10px", border: `1px solid ${COLORS.cardBorder}`, background: isProcessed ? "#FAFAF8" : "#FFFFFF" }}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: "10px" }}>
        <div>
          <span style={{ display: "flex", alignItems: "center", gap: "6px", fontWeight: 700, fontSize: "13.5px", color: COLORS.navy }}>
            <ClipboardList size={14} color={COLORS.gold} />{request.sujet || "Demande"}
          </span>
          <span style={{ display: "block", fontSize: "12px", color: COLORS.textMuted, marginTop: "2px" }}>
            {formatRequestDate(request.createdAt)}{request.produit ? ` — ${request.produit}` : ""}
          </span>
        </div>
        <span style={{
          fontSize: "11px", fontWeight: 700, padding: "4px 10px", borderRadius: "999px", whiteSpace: "nowrap",
          background: isProcessed ? "#EAF7EE" : "#FFF4DE", color: isProcessed ? "#1E7A34" : COLORS.goldDark,
        }}>
          {isProcessed ? "Traitée" : "En attente"}
        </span>
      </div>
      <div style={{ marginTop: "10px", fontSize: "13px", color: COLORS.navy, display: "flex", flexDirection: "column", gap: "3px" }}>
        {isAdmin && <span><strong>Demandeur :</strong> {request.prenom} {request.nom} ({request.requestedBy || request.email})</span>}
        {request.machine && <span><strong>Machine :</strong> {request.machine}{request.numeroSerie ? ` — N° série ${request.numeroSerie}` : ""}</span>}
        {request.categorie && <span><strong>Catégorie :</strong> {request.categorie}</span>}
        {request.motif && <span><strong>Motif :</strong> {request.motif}</span>}
        {request.messageHtml ? (
          <div style={{ color: COLORS.textMuted, marginTop: "4px", lineHeight: 1.5 }} dangerouslySetInnerHTML={{ __html: makeImagesEmbeddable(request.messageHtml) }} />
        ) : (
          request.message && <span style={{ color: COLORS.textMuted, marginTop: "4px" }}>{request.message}</span>
        )}
      </div>
      {request.pieceJointeUrl && (
        <a href={request.pieceJointeUrl} target="_blank" rel="noreferrer" style={{ display: "inline-flex", alignItems: "center", gap: "6px", marginTop: "10px", fontSize: "12.5px", fontWeight: 700, color: COLORS.goldDark, textDecoration: "none" }}>
          <Paperclip size={13} />{request.pieceJointeNom || "Pièce jointe"}<Download size={13} />
        </a>
      )}
      {!request.messageHtml && Array.isArray(request.imagesMessage) && request.imagesMessage.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "10px" }}>
          {request.imagesMessage.map((img, i) => (
            <a key={img.url || i} href={img.url} target="_blank" rel="noreferrer" title={img.nom || "Image"}>
              <img src={driveEmbeddableUrl(img.url)} alt={img.nom || "Image jointe"} style={{ width: "64px", height: "64px", objectFit: "cover", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}` }} />
            </a>
          ))}
        </div>
      )}
      {isProcessed && (
        <>
          {(request.reponseHtml || request.reponse) && (
            <div style={{ marginTop: "10px", padding: "10px 12px", borderRadius: "8px", background: "#FFFFFF", border: `1px solid ${COLORS.cardBorder}` }}>
              <span style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", color: COLORS.textMuted }}>Réponse envoyée</span>
              {request.reponseHtml ? (
                <div style={{ fontSize: "13px", color: COLORS.navy, marginTop: "4px", lineHeight: 1.5 }} dangerouslySetInnerHTML={{ __html: makeImagesEmbeddable(request.reponseHtml) }} />
              ) : (
                <p style={{ fontSize: "13px", color: COLORS.navy, marginTop: "4px", whiteSpace: "pre-wrap" }}>{request.reponse}</p>
              )}
              {request.reponsePieceJointeUrl && (
                <a href={request.reponsePieceJointeUrl} target="_blank" rel="noreferrer" style={{ display: "inline-flex", alignItems: "center", gap: "6px", marginTop: "8px", fontSize: "12.5px", fontWeight: 700, color: COLORS.goldDark, textDecoration: "none" }}>
                  <Paperclip size={13} />{request.reponsePieceJointeNom || "Pièce jointe"}<Download size={13} />
                </a>
              )}
              {!request.reponseHtml && Array.isArray(request.reponseImages) && request.reponseImages.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "8px" }}>
                  {request.reponseImages.map((img, i) => (
                    <a key={img.url || i} href={img.url} target="_blank" rel="noreferrer" title={img.nom || "Image"}>
                      <img src={driveEmbeddableUrl(img.url)} alt={img.nom || "Image jointe"} style={{ width: "64px", height: "64px", objectFit: "cover", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}` }} />
                    </a>
                  ))}
                </div>
              )}
            </div>
          )}
          {request.processedAt && (
            <p style={{ fontSize: "11.5px", color: COLORS.textMuted, marginTop: "8px" }}>
              Traitée le {formatRequestDate(request.processedAt)}{request.processedBy ? ` par ${request.processedBy}` : ""}
            </p>
          )}
        </>
      )}
      {isAdmin && !isProcessed && (
        <div style={{ marginTop: "12px", display: "flex", flexDirection: "column", gap: "8px" }}>
          <label style={{ fontSize: "11.5px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", color: COLORS.textMuted }}>
            Réponse au demandeur
          </label>
          <div>
            <style>{`
              .reponse-message-editor:empty:before {
                content: attr(data-placeholder);
                color: ${COLORS.textMuted};
                pointer-events: none;
              }
              .reponse-message-editor img:hover {
                outline: 2px solid ${COLORS.gold};
              }
            `}</style>
            <div
              ref={reponseEditorRef}
              className="reponse-message-editor"
              contentEditable
              suppressContentEditableWarning
              onPaste={handleReponsePaste}
              onInput={syncReponseState}
              onDoubleClick={handleReponseEditorDoubleClick}
              data-placeholder="Écrivez votre réponse, elle sera envoyée par e-mail au demandeur…"
              style={{ ...formInputStyle, minHeight: "180px", maxHeight: "480px", overflowY: "auto", fontFamily: "inherit", lineHeight: 1.5, cursor: "text" }}
            />
            <span style={{ fontSize: "11.5px", color: COLORS.textMuted, fontStyle: "italic", display: "block", marginTop: "4px" }}>
              Collez une capture d'écran (Ctrl+V) directement dans la réponse ci-dessus
            </span>
            {reponseImageCount > 0 && (
              <p style={{ fontSize: "11.5px", color: COLORS.textMuted, marginTop: "4px" }}>
                {reponseImageCount} image{reponseImageCount > 1 ? "s" : ""} insérée{reponseImageCount > 1 ? "s" : ""} — double-cliquez sur une image pour la retirer.
              </p>
            )}
          </div>
          <div>
            <label style={{ fontSize: "11.5px", color: COLORS.textMuted, display: "flex", alignItems: "center", gap: "6px", marginBottom: "6px" }}>
              <Paperclip size={13} />Pièce jointe (optionnel)
            </label>
            <input type="file" onChange={(e) => setReponseFile(e.target.files?.[0] || null)} style={{ fontSize: "13px" }} />
          </div>
          <button
            onClick={handleSend}
            disabled={archiving || !canSend}
            title={!canSend ? "Écrivez une réponse avant de traiter la demande" : undefined}
            style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "6px", background: COLORS.navy, color: "#FFFFFF", border: "none", borderRadius: "8px", padding: "10px 14px", fontSize: "12.5px", fontWeight: 700, cursor: canSend ? "pointer" : "not-allowed", opacity: archiving || !canSend ? 0.6 : 1 }}
          >
            <Send size={14} />{archiving ? "Envoi…" : "Envoyer la réponse et marquer comme traitée"}
          </button>
        </div>
      )}
    </div>
  );
}

function ExpertRequestsModal({ requests, isAdmin, currentUserEmail, onClose }) {
  const [tab, setTab] = useState("active"); // "active" | "history"
  const [archivingId, setArchivingId] = useState(null);
  const [error, setError] = useState(null);
  const [clearing, setClearing] = useState(false);

  const activeRequests = requests.filter((r) => !r.archived);
  const archivedRequests = requests.filter((r) => r.archived);
  const list = tab === "active" ? activeRequests : archivedRequests;

  const handleClearHistory = async () => {
    if (archivedRequests.length === 0) return;
    const confirmed = window.confirm(
      `Effacer définitivement les ${archivedRequests.length} demande(s) de l'historique ? Cette action est irréversible.`
    );
    if (!confirmed) return;
    setError(null);
    setClearing(true);
    try {
      await Promise.all(archivedRequests.map((r) => deleteDoc(doc(db, "expertRequests", r.id))));
    } catch {
      setError("Impossible d'effacer l'historique pour l'instant. Réessayez.");
    } finally {
      setClearing(false);
    }
  };

  const handleArchive = async (request, reponse, file, images = [], html = "") => {
    setError(null);
    setArchivingId(request.id);
    try {
      let reponsePieceJointeUrl = null;
      if (file) {
        try {
          reponsePieceJointeUrl = await uploadFileToDrive(file);
        } catch (err) {
          console.error("Erreur upload pièce jointe réponse:", err);
          setError("La pièce jointe n'a pas pu être envoyée. La réponse va tout de même être envoyée sans elle.");
        }
      }
      let reponseImages = [];
      let reponseHtml = html;
      if (images.length) {
        const tempDiv = document.createElement("div");
        tempDiv.innerHTML = html;
        const imgEls = Array.from(tempDiv.querySelectorAll("img[data-image-id]"));
        let uploadFailed = false;
        // Chaque image est uploadée indépendamment : si l'une échoue, les
        // autres sont quand même enregistrées. Une image en échec est
        // retirée du HTML plutôt que laissée avec son URL blob: locale
        // (invalide dès que la page est rechargée ou vue ailleurs, ce qui
        // faisait "disparaître" l'image après l'envoi de la réponse).
        await Promise.all(imgEls.map(async (imgEl) => {
          const id = imgEl.getAttribute("data-image-id");
          const entry = images.find((img) => img.id === id);
          if (!entry?.file) { imgEl.remove(); return; }
          try {
            const url = await uploadFileToDrive(entry.file);
            reponseImages.push({ nom: entry.file.name, url });
            imgEl.src = url;
            imgEl.removeAttribute("data-image-id");
            imgEl.removeAttribute("title");
          } catch (err) {
            console.error("Erreur upload image réponse:", err);
            uploadFailed = true;
            imgEl.remove();
          }
        }));
        reponseHtml = tempDiv.innerHTML;
        if (uploadFailed) {
          setError("Une ou plusieurs images n'ont pas pu être envoyées. La réponse va tout de même être envoyée sans elles.");
        }
      }
      await updateDoc(doc(db, "expertRequests", request.id), {
        archived: true,
        status: "Traitée",
        reponse,
        reponseHtml,
        reponsePieceJointeUrl: reponsePieceJointeUrl || null,
        reponsePieceJointeNom: reponsePieceJointeUrl ? file?.name || null : null,
        reponseImages,
        processedAt: serverTimestamp(),
        processedBy: currentUserEmail,
      });
      try {
        await fetch(APPS_SCRIPT_URL, {
          method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            type: "reponse_demande",
            requestId: request.id,
            email: request.email,
            nom: request.nom,
            prenom: request.prenom,
            sujet: request.sujet,
            produit: request.produit || "",
            machine: request.machine || "",
            messageOriginal: [request.motif ? `Motif : ${request.motif}` : "", request.message || ""].filter(Boolean).join("\n\n"),
            reponse,
            pieceJointe: reponsePieceJointeUrl ? file?.name || "Pièce jointe" : "Aucune",
            imagesReponse: reponseImages.map((i) => i.nom).join(", ") || "Aucune",
            traitePar: currentUserEmail,
          }),
        });
      } catch {
        // La demande est bien marquée comme traitée même si l'envoi de l'e-mail échoue ;
        // on informe simplement l'admin que la notification n'est pas partie.
        setError("La demande a été traitée mais l'e-mail de réponse n'a pas pu être envoyé.");
      }
    } catch (err) {
      console.error("Erreur handleArchive:", err);
      setError(`Impossible de mettre à jour la demande pour l'instant (${err?.code || err?.message || "erreur inconnue"}). Réessayez.`);
    } finally {
      setArchivingId(null);
    }
  };

  const tabBtnStyle = (active) => ({
    display: "flex", alignItems: "center", gap: "6px", flex: 1, justifyContent: "center",
    padding: "9px 12px", borderRadius: "8px", border: "none", cursor: "pointer",
    fontSize: "12.5px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em",
    background: active ? COLORS.gold : "transparent", color: active ? COLORS.navy : COLORS.textMuted,
  });

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: "20px" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#FFFFFF", borderRadius: "16px", maxWidth: "900px", width: "100%", maxHeight: "92vh", overflow: "auto", padding: "28px", position: "relative" }}>
        <button onClick={onClose} aria-label="Fermer" style={{ position: "absolute", top: "16px", right: "16px", background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={20} /></button>
        <h2 style={{ fontSize: "20px", fontWeight: 800, marginBottom: "6px", display: "flex", alignItems: "center", gap: "8px" }}>
          <ClipboardList size={20} color={COLORS.gold} />Demandes aux experts
        </h2>
        <p style={{ fontSize: "13px", color: COLORS.textMuted, marginBottom: "16px" }}>
          {isAdmin ? "Gérez les demandes en cours et consultez l'historique des demandes traitées." : "Suivez l'état de vos demandes et consultez votre historique."}
        </p>
        <div style={{ display: "flex", gap: "8px", background: "#FAFAF8", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "10px", padding: "4px", marginBottom: "16px" }}>
          <button onClick={() => setTab("active")} style={tabBtnStyle(tab === "active")}>
            <Inbox size={14} />En cours ({activeRequests.length})
          </button>
          <button onClick={() => setTab("history")} style={tabBtnStyle(tab === "history")}>
            <History size={14} />Historique ({archivedRequests.length})
          </button>
        </div>
        {isAdmin && tab === "history" && archivedRequests.length > 0 && (
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: "12px" }}>
            <button
              onClick={handleClearHistory}
              disabled={clearing}
              style={{
                display: "flex", alignItems: "center", gap: "6px", background: "none", border: `1px solid #C0392B`,
                color: "#C0392B", borderRadius: "8px", padding: "6px 12px", fontSize: "12px", fontWeight: 700,
                cursor: clearing ? "not-allowed" : "pointer", opacity: clearing ? 0.6 : 1,
              }}
            >
              <Trash2 size={13} />{clearing ? "Suppression…" : "Effacer l'historique"}
            </button>
          </div>
        )}
        {error && <p style={{ color: "#C0392B", fontSize: "13px", marginBottom: "10px" }}>{error}</p>}
        {list.length === 0 ? (
          <p style={{ fontSize: "14px", color: COLORS.textMuted, textAlign: "center", padding: "24px 0" }}>
            {tab === "active" ? "Aucune demande en cours." : "Aucune demande archivée pour le moment."}
          </p>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
            {list.map((r) => (
              <ExpertRequestCard key={r.id} request={r} isAdmin={isAdmin} onArchive={handleArchive} archiving={archivingId === r.id} />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// Retrouve le nœud produit (ou sous-produit) correspondant à un id, dans une catégorie donnée.
function findProductNode(categories, categoryId, productId) {
  const category = categories.find((c) => c.id === categoryId);
  if (!category) return null;
  for (const item of category.items || []) {
    if (item.id === productId) return item;
    if (item.items) {
      const sub = item.items.find((s) => s.id === productId);
      if (sub) return sub;
    }
  }
  return null;
}

const adminLabelStyle = { fontSize: "12px", fontWeight: 700, color: COLORS.textMuted, marginBottom: "4px", display: "block", textTransform: "uppercase", letterSpacing: "0.02em" };

// Fenêtre d'administration : envoie un PDF directement dans le dossier Drive
// du produit et du type de document choisis (mêmes IDs que ceux définis dans CATEGORIES).
function AdminUploadModal({ categories, onClose }) {
  const [selectedCategory, setSelectedCategory] = useState("");
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedFolderType, setSelectedFolderType] = useState("");
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [message, setMessage] = useState(null);
  const [folderFiles, setFolderFiles] = useState([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [deletingId, setDeletingId] = useState(null);

  const category = categories.find((c) => c.id === selectedCategory) || null;
  // Garantie, Pièces détachées et Formation n'ont pas d'arborescence Produit :
  // leurs documents vivent directement dans le dossier Drive de la catégorie.
  const isCategoryLevelDoc = category ? ["garantie", "pieces-detachees", "formation"].includes(category.id) : false;
  const productNode = !isCategoryLevelDoc && category ? findProductNode(categories, selectedCategory, selectedProduct) : null;
  const availableFolderTypes = productNode
    ? Object.entries(DRIVE_FOLDER_LABELS).filter(([key]) => productNode[key])
    : [];
  const currentFolderId = isCategoryLevelDoc
    ? category?.driveFolderId || null
    : (productNode && selectedFolderType ? productNode[selectedFolderType] : null);

  // Charge la liste des fichiers du dossier Drive dès qu'un type de document
  // (donc un folderId) est choisi, pour permettre de les supprimer.
  useEffect(() => {
    if (!currentFolderId) { setFolderFiles([]); return; }
    let cancelled = false;
    setLoadingFiles(true);
    getDriveFiles(currentFolderId)
      .then((files) => { if (!cancelled) setFolderFiles(files); })
      .catch(() => { if (!cancelled) setFolderFiles([]); })
      .finally(() => { if (!cancelled) setLoadingFiles(false); });
    return () => { cancelled = true; };
  }, [currentFolderId]);

  const handleDeleteFile = async (fileId) => {
    if (!window.confirm("Supprimer définitivement ce fichier du Drive ?")) return;
    setDeletingId(fileId);
    setMessage(null);
    try {
      await deleteFileFromDrive(fileId);
      setFolderFiles((prev) => prev.filter((f) => f.id !== fileId));
      setMessage({ type: "success", text: "Fichier supprimé." });
    } catch (err) {
      setMessage({ type: "error", text: err.message || "Échec de la suppression." });
    } finally {
      setDeletingId(null);
    }
  };

  const handleCategoryChange = (id) => {
    setSelectedCategory(id);
    setSelectedProduct("");
    setSelectedFolderType("");
    setMessage(null);
  };

  const handleProductChange = (id) => {
    setSelectedProduct(id);
    setSelectedFolderType("");
    setMessage(null);
  };

  const handleUpload = async (e) => {
    e.preventDefault();
    if (!file || !currentFolderId) return;
    const folderId = currentFolderId;
    if (!folderId) {
      setMessage({ type: "error", text: "Aucun dossier Drive associé à ce type de document pour ce produit." });
      return;
    }
    setUploading(true);
    setUploadProgress(0);
    setMessage(null);
    try {
      await uploadFileToDrive(file, folderId, setUploadProgress);
      setMessage({ type: "success", text: "Fichier envoyé avec succès sur Google Drive." });
      setFile(null);
      getDriveFiles(folderId).then(setFolderFiles).catch(() => {});
    } catch (err) {
      setMessage({ type: "error", text: err.message || "Échec de l'envoi." });
    } finally {
      setUploading(false);
      setUploadProgress(0);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,31,58,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: "20px" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#FFFFFF", borderRadius: "16px", maxWidth: "460px", width: "100%", maxHeight: "85vh", overflow: "auto", padding: "28px", position: "relative" }}>
        <button onClick={onClose} aria-label="Fermer" style={{ position: "absolute", top: "16px", right: "16px", background: "none", border: "none", cursor: "pointer", color: COLORS.textMuted }}><X size={20} /></button>
        <h2 style={{ fontSize: "20px", fontWeight: 800, marginBottom: "6px", display: "flex", alignItems: "center", gap: "8px" }}>
          <Upload size={20} color={COLORS.gold} />Ajouter / Supprimer un document
        </h2>
        <p style={{ fontSize: "13px", color: COLORS.textMuted, marginBottom: "18px" }}>
          Envoyez un fichier directement dans le dossier Drive du produit concerné.
        </p>

        <form onSubmit={handleUpload} style={{ display: "flex", flexDirection: "column", gap: "14px" }}>
          <div>
            <label style={adminLabelStyle}>Catégorie</label>
            <select value={selectedCategory} onChange={(e) => handleCategoryChange(e.target.value)} style={{ ...selectStyle, width: "100%" }}>
              <option value="">-- Choisir --</option>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
            </select>
          </div>

          {!isCategoryLevelDoc && (
            <div>
              <label style={adminLabelStyle}>Produit</label>
              <select value={selectedProduct} onChange={(e) => handleProductChange(e.target.value)} disabled={!category} style={{ ...selectStyle, width: "100%" }}>
                <option value="">-- Choisir --</option>
                {category?.items?.map((p) =>
                  p.items ? (
                    <optgroup label={p.label} key={p.id}>
                      {p.items.map((sub) => <option key={sub.id} value={sub.id}>{sub.label}</option>)}
                    </optgroup>
                  ) : (
                    <option key={p.id} value={p.id}>{p.label}</option>
                  )
                )}
              </select>
            </div>
          )}

          {!isCategoryLevelDoc && (
            <div>
              <label style={adminLabelStyle}>Type de document</label>
              <select value={selectedFolderType} onChange={(e) => setSelectedFolderType(e.target.value)} disabled={!productNode} style={{ ...selectStyle, width: "100%" }}>
                <option value="">-- Choisir --</option>
                {availableFolderTypes.map(([key, label]) => <option key={key} value={key}>{label}</option>)}
              </select>
              {productNode && availableFolderTypes.length === 0 && (
                <p style={{ fontSize: "12px", color: COLORS.textMuted, marginTop: "4px" }}>Aucun dossier Drive configuré pour ce produit.</p>
              )}
            </div>
          )}

          {isCategoryLevelDoc && category && !category.driveFolderId && (
            <p style={{ fontSize: "12px", color: "#C0392B", margin: 0 }}>
              Aucun dossier Drive n'est encore configuré pour « {category.label} ». Ajoutez son id dans CATEGORIES (champ driveFolderId) pour activer l'upload et la suppression ici.
            </p>
          )}

          {currentFolderId && (
            <div>
              <label style={adminLabelStyle}>Fichiers dans ce dossier</label>
              {loadingFiles ? (
                <p style={{ fontSize: "13px", color: COLORS.textMuted, margin: 0 }}>Chargement…</p>
              ) : folderFiles.length === 0 ? (
                <p style={{ fontSize: "13px", color: COLORS.textMuted, margin: 0 }}>Aucun fichier dans ce dossier.</p>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: "6px", maxHeight: "160px", overflow: "auto", border: `1px solid ${COLORS.cardBorder}`, borderRadius: "8px", padding: "6px" }}>
                  {folderFiles.map((f) => (
                    <div key={f.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px", padding: "4px 6px" }}>
                      <a href={f.url} target="_blank" rel="noreferrer" title={f.name} style={{ fontSize: "13px", color: COLORS.navy, textDecoration: "none", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</a>
                      <button
                        type="button"
                        onClick={() => handleDeleteFile(f.id)}
                        disabled={deletingId === f.id}
                        aria-label={`Supprimer ${f.name}`}
                        style={{ display: "flex", alignItems: "center", justifyContent: "center", background: "none", border: "none", color: "#C0392B", cursor: deletingId === f.id ? "not-allowed" : "pointer", flexShrink: 0, opacity: deletingId === f.id ? 0.5 : 1 }}
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          <div>
            <label style={adminLabelStyle}>Fichier</label>
            <label style={{ display: "flex", alignItems: "center", gap: "10px", width: "100%", padding: "8px 12px", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}`, background: "#FFFFFF", cursor: "pointer", boxSizing: "border-box" }}>
              <span style={{ flexShrink: 0, padding: "6px 12px", borderRadius: "6px", border: `1px solid ${COLORS.cardBorder}`, background: "#F5F5F5", fontSize: "12.5px", fontWeight: 600, color: COLORS.navy }}>Choisir un fichier</span>
              <span style={{ fontSize: "13px", color: file ? COLORS.navy : COLORS.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{file ? file.name : "Aucun fichier choisi"}</span>
              <input type="file" onChange={(e) => setFile(e.target.files[0] || null)} style={{ position: "absolute", width: "1px", height: "1px", padding: 0, margin: "-1px", overflow: "hidden", clip: "rect(0,0,0,0)", border: 0 }} />
            </label>
          </div>

          {uploading && (
            <div>
              <div style={{ height: "6px", borderRadius: "3px", background: "rgba(0,0,0,0.08)", overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${uploadProgress}%`, background: COLORS.gold, transition: "width 0.2s ease" }} />
              </div>
              <p style={{ fontSize: "12px", color: COLORS.textMuted, marginTop: "4px", marginBottom: 0 }}>Envoi en cours… {uploadProgress}%</p>
            </div>
          )}

          {message && (
            <p style={{ fontSize: "13px", color: message.type === "error" ? "#C0392B" : "#1E7A34", margin: 0 }}>{message.text}</p>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px", marginTop: "4px" }}>
            <button type="button" onClick={onClose} style={{ padding: "10px 18px", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}`, background: "#FFFFFF", cursor: "pointer", fontSize: "13px", fontWeight: 600 }}>Fermer</button>
            <button
              type="submit"
              disabled={uploading || !file || !currentFolderId}
              style={{ padding: "10px 18px", borderRadius: "8px", border: "none", background: COLORS.gold, color: COLORS.navy, cursor: uploading ? "not-allowed" : "pointer", fontSize: "13px", fontWeight: 700, opacity: uploading || !file || !currentFolderId ? 0.6 : 1 }}
            >
              {uploading ? "Envoi…" : "Uploader"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function LoginScreen({ kickedOut }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [resetSent, setResetSent] = useState(false);
  const handleSubmit = async (e) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try { await signInWithEmailAndPassword(auth, email, password); }
    catch (err) { setError(err.code === "auth/invalid-credential" || err.code === "auth/wrong-password" ? "E-mail ou mot de passe incorrect." : "Impossible de se connecter."); }
    finally { setLoading(false); }
  };
  const handleForgotPassword = async () => {
    if (!email) { setError("Entrez votre e-mail."); return; }
    try { await sendPasswordResetEmail(auth, email); setResetSent(true); } catch { setError("Impossible d'envoyer l'e-mail."); }
  };
  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: "20px", background: COLORS.bg, padding: "20px" }}>
      <img src={logo} alt="Kaeser Compresseurs" style={{ height: "70px", width: "auto" }} />
      {kickedOut && <p style={alertStyle}>Vous avez été déconnecté(e) car ce compte a été utilisé sur un autre appareil.</p>}
      {resetSent && <p style={{ ...alertStyle, background: "#EAF7EE", color: "#1E7A34" }}>E-mail envoyé.</p>}
      {error && <p style={alertStyle}>{error}</p>}
      <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: "10px", width: "100%", maxWidth: "320px" }}>
        <input type="email" placeholder="votre.email@kaeser.com" value={email} onChange={(e) => setEmail(e.target.value)} required style={inputStyle} />
        <input type="password" placeholder="Mot de passe" value={password} onChange={(e) => setPassword(e.target.value)} required style={inputStyle} />
        <button type="submit" disabled={loading} style={{ background: COLORS.gold, color: COLORS.navy, border: "none", borderRadius: "10px", padding: "14px 28px", fontSize: "14px", fontWeight: 700, cursor: "pointer" }}>{loading ? "Connexion…" : "Se connecter"}</button>
      </form>
      <button onClick={handleForgotPassword} style={{ background: "none", border: "none", color: COLORS.textMuted, fontSize: "13px", cursor: "pointer", textDecoration: "underline" }}>Mot de passe oublié</button>
    </div>
  );
}

const alertStyle = { background: "#FBE9E7", color: "#C0392B", fontSize: "13px", padding: "10px 16px", borderRadius: "8px", maxWidth: "320px", textAlign: "center" };
const inputStyle = { padding: "12px 14px", borderRadius: "8px", border: `1px solid ${COLORS.cardBorder}`, fontSize: "14px", outline: "none" };
const gridStyle = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 240px))", gap: "18px", maxWidth: "1560px", margin: "0 auto", justifyContent: "center" };
const searchWrapStyle = { display: "flex", alignItems: "center", gap: "10px", maxWidth: "480px", marginLeft: "auto", marginRight: "auto", marginBottom: "28px", padding: "12px 16px", borderRadius: "10px", border: `1px solid ${COLORS.cardBorder}`, background: "#FFFFFF", boxSizing: "border-box" };
const searchInputStyle = { flex: 1, minWidth: 0, border: "none", outline: "none", fontSize: "14px", color: COLORS.navy, background: "transparent" };
const selectStyle = { width: "100%", maxWidth: "320px", minWidth: 0, flexShrink: 1, padding: "0 14px", height: "42px", borderRadius: "10px", border: `1px solid ${COLORS.cardBorder}`, background: "#FFFFFF", color: COLORS.navy, fontSize: "13px", boxSizing: "border-box" };
const navBtnStyle = { display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(255,255,255,0.15)", border: "none", color: "#FFFFFF", cursor: "pointer", padding: "6px", borderRadius: "6px" };
const closeBtnStyle = { display: "flex", alignItems: "center", gap: "6px", background: "rgba(255,255,255,0.15)", border: "none", color: "#FFFFFF", fontSize: "13px", cursor: "pointer", padding: "6px 12px", borderRadius: "8px" };