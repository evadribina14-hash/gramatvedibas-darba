import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  BriefcaseBusiness,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  ClipboardCheck,
  Clock3,
  Download,
  FileText,
  Info,
  Plus,
  Save,
  Settings2,
  Trash2,
  Users,
  X,
} from "lucide-react";
import "./styles.css";

const STORAGE = "gramatvedibas-darba-v4";

const uid = (prefix = "id") =>
  `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

const todayISO = () => new Date().toISOString().slice(0, 10);

const defaultInfo = [
  {
    id: "iin",
    category: "Nodokļu likmes",
    title: "IIN",
    value:
      "2026. gadā IIN likme 25,5% gada ienākuma daļai līdz 105 300 EUR. Virs 105 300 EUR – 33%. Papildu 3% piemēro gada ienākumu pārsniegumam virs 200 000 EUR.",
    year: "2026",
  },
  {
    id: "neapliekamais",
    category: "Nodokļu likmes",
    title: "Neapliekamais minimums",
    value: "2026. gadā – 550 EUR mēnesī.",
    year: "2026",
  },
  {
    id: "apgadamais",
    category: "Nodokļu likmes",
    title: "Atvieglojums par apgādājamo",
    value: "250 EUR mēnesī par vienu apgādājamo.",
    year: "2026",
  },
  {
    id: "pensionars",
    category: "Nodokļu likmes",
    title: "Pensionāra neapliekamais minimums",
    value: "1000 EUR mēnesī jeb 12 000 EUR gadā.",
    year: "2026",
  },
  {
    id: "vsao",
    category: "Nodokļu likmes",
    title: "VSAOI – vispārējā likme",
    value:
      "Darba ņēmējs 10,50%. Darba devējs 23,59%. Kopā 34,09%.",
    year: "2026",
  },
  {
    id: "vsao-pensionars",
    category: "Nodokļu likmes",
    title: "VSAOI – pensionārs",
    value:
      "Darba ņēmējs 9,25%. Darba devējs 20,77%. Kopā 30,02%.",
    year: "2026",
  },
  {
    id: "pvn",
    category: "PVN",
    title: "PVN standarta likme",
    value: "21%.",
    year: "2026",
  },
  {
    id: "pvn-reg",
    category: "PVN",
    title: "PVN reģistrācijas slieksnis",
    value: "50 000 EUR.",
    year: "2026",
  },
  {
    id: "pvn-eu",
    category: "PVN",
    title: "ES preču iegādes slieksnis",
    value: "10 000 EUR.",
    year: "2026",
  },
  {
    id: "uznemumu",
    category: "Uzņēmumi",
    title: "Mikrosabiedrība",
    value:
      "Bilances kopsumma līdz 450 000 EUR; neto apgrozījums līdz 900 000 EUR; vidējais darbinieku skaits līdz 10.",
    year: "2026",
  },
  {
    id: "uznemumu-small",
    category: "Uzņēmumi",
    title: "Maza sabiedrība",
    value:
      "Bilances kopsumma līdz 5 000 000 EUR; net
