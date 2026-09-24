import React, {useEffect, useMemo, useState} from "react";
import {createRoot} from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Clock3, Table2, CalendarDays, BookOpen, ClipboardCheck, Bell,
  Plus, Play, Square, Pause, Download, ChevronLeft, ChevronRight,
  Check, Trash2, Save, Settings2
} from "lucide-react";
import "./styles.css";

const STORAGE = "gramatvedibas-darba-v1";

const defaultData = {
  clients: [
    {id:"c1", name:"SIA ABC", vat:"monthly"},
    {id:"c2", name:"SIA XYZ", vat:"quarterly"},
    {id:"c3", name:"IK Anna", vat:"none"}
  ],
  jobs: [
    "Grāmatojumi",
    "PVN deklarācija",
    "Algu aprēķins",
    "Bankas kontrole",
    "Salīdzināšanas akti"
  ],
  entries: [],
  reports: {},
  annual: {},
  notes: {}
};

const
