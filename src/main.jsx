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
      "Bilances kopsumma līdz 5 000 000 EUR; neto apgrozījums līdz 10 000 000 EUR; vidējais darbinieku skaits līdz 50.",
    year: "2026",
  },
  {
    id: "uznemumu-medium",
    category: "Uzņēmumi",
    title: "Vidēja sabiedrība",
    value:
      "Bilances kopsumma līdz 25 000 000 EUR; neto apgrozījums līdz 50 000 000 EUR; vidējais darbinieku skaits līdz 250.",
    year: "2026",
  },
  {
    id: "komandejumi",
    category: "Komandējumi",
    title: "Komandējumu dienas nauda Latvijā",
    value:
      "Latvijas teritorijā dienas naudas norma – 8 EUR. Ārvalstu normas jāskatās pēc konkrētās valsts.",
    year: "2026",
  },
  {
    id: "darba-nem",
    category: "Darba ņēmēji",
    title: "Ziņas par darba ņēmējiem",
    value:
      "Darba ņēmēju ziņas iesniedzamas atbilstoši VID prasībām un MK noteikumiem Nr. 827.",
    year: "2026",
  },
  {
    id: "pamatl",
    category: "Grāmatvedība",
    title: "Pamatlīdzekļi",
    value:
      "Pārbaudīt iegādi, nodošanu ekspluatācijā, inventāra numurus, nolietojumu, atlikumus un inventarizāciju.",
    year: "2026",
  },
  {
    id: "reprezentacija",
    category: "Grāmatvedība",
    title: "Reprezentācijas izdevumi",
    value:
      "Pārbaudīt dokumentus, izdevumu ekonomisko būtību un nodokļu piemērošanu.",
    year: "2026",
  },
];

const standardAnnualTasks = [
  {
    group: "Bilance un konti",
    tasks: [
      "2310 atlikums pārbaudīts",
      "2310 pareizā pusē",
      "2380 atlikums pārbaudīts",
      "Bankas atlikums pārbaudīts",
      "Kases atlikums pārbaudīts",
    ],
  },
  {
    group: "Debitori un kreditori",
    tasks: [
      "Nosūtīti salīdzināšanas akti",
      "Saņemti salīdzināšanas akti",
      "Debitoru atlikumi pārbaudīti",
      "Kreditoru atlikumi pārbaudīti",
    ],
  },
  {
    group: "Pamatlīdzekļi",
    tasks: [
      "Veikta pamatlīdzekļu inventarizācija",
      "Pamatlīdzekļu atlikumi pārbaudīti",
      "Nolietojums pārbaudīts",
    ],
  },
  {
    group: "Ieņēmumi un izdevumi",
    tasks: [
      "Ieņēmumi pārbaudīti",
      "Izdevumi pārbaudīti",
      "Uzkrājumi pārbaudīti",
    ],
  },
];

const defaultReportTypes = [
  {
    id: "vat",
    name: "PVN deklarācija",
    period: "monthly",
    dueDay: 20,
    builtIn: true,
  },
  {
    id: "ddz",
    name: "Darba devēja ziņojums",
    period: "monthly",
    dueDay: 17,
    builtIn: true,
  },
  {
    id: "uin",
    name: "UIN deklarācija",
    period: "monthly",
    dueDay: 20,
    builtIn: true,
  },
];

/*
  SVARĪGI:
  Šajā jaunajā versijā sākotnējais klientu saraksts ir TUKŠS.
  Tātad vecie testa klienti no iepriekšējās versijas netiks
  izmantoti kā jaunie sākotnējie klienti.

  Ja vecā versija jau bija saglabājusi datus telefonā,
  jaunais STORAGE nosaukums izveido jaunu datu versiju.
*/

const defaultData = {
  clients: [],
  jobs: [
    {
      id: "job-1",
      name: "Ikdienas grāmatvedība",
    },
    {
      id: "job-2",
      name: "PVN deklarācija",
    },
    {
      id: "job-3",
      name: "Darba devēja ziņojums",
    },
  ],
  entries: [],
  notes: [],
  info: defaultInfo,
  annual: {},
  reportTypes: defaultReportTypes,
  reportStatuses: {},
};

function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE);

    if (!saved) {
      return defaultData;
    }

    const parsed = JSON.parse(saved);

    return {
      ...defaultData,
      ...parsed,
      clients: Array.isArray(parsed.clients)
        ? parsed.clients
        : [],
      jobs: Array.isArray(parsed.jobs)
        ? parsed.jobs
        : defaultData.jobs,
      entries: Array.isArray(parsed.entries)
        ? parsed.entries
        : [],
      notes: Array.isArray(parsed.notes)
        ? parsed.notes
        : [],
      info: Array.isArray(parsed.info)
        ? parsed.info
        : defaultInfo,
      annual: parsed.annual || {},
      reportTypes: Array.isArray(parsed.reportTypes)
        ? parsed.reportTypes
        : defaultReportTypes,
      reportStatuses: parsed.reportStatuses || {},
    };
  } catch {
    return defaultData;
  }
}

function formatDate(date) {
  if (!date) return "";

  const d = new Date(`${date}T00:00:00`);

  return d.toLocaleDateString("lv-LV");
}

function formatDuration(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;

  if (h === 0) return `${m} min`;

  if (m === 0) return `${h} h`;

  return `${h} h ${m} min`;
}

function minutesToHours(minutes) {
  return minutes / 60;
}

function roundHours(minutes) {
  return Math.round(minutes / 60);
}

function monthDates(month) {
  const [year, monthNumber] = month.split("-").map(Number);

  const lastDay = new Date(
    year,
    monthNumber,
    0
  ).getDate();

  return Array.from(
    { length: lastDay },
    (_, i) => {
      const day = i + 1;

      return `${year}-${String(monthNumber).padStart(
        2,
        "0"
      )}-${String(day).padStart(2, "0")}`;
    }
  );
}

function monthLabel(month) {
  const [year, m] = month.split("-").map(Number);

  return new Date(year, m - 1, 1).toLocaleDateString(
    "lv-LV",
    {
      month: "long",
      year: "numeric",
    }
  );
}

function App() {
  const [data, setData] = useState(loadData);

  const [page, setPage] = useState("work");

  const [selectedMonth, setSelectedMonth] =
    useState(
      new Date()
        .toISOString()
        .slice(0, 7)
    );

  const [clientId, setClientId] = useState("");

  const [jobId, setJobId] = useState("");

  const [note, setNote] = useState("");

  const [running, setRunning] = useState(null);

  const [now, setNow] = useState(Date.now());

  const [newClient, setNewClient] = useState("");

  const [newJob, setNewJob] = useState("");

  const [infoCategory, setInfoCategory] =
    useState("");

  const [infoTitle, setInfoTitle] =
    useState("");

  const [infoValue, setInfoValue] =
    useState("");

  const [infoYear, setInfoYear] =
    useState("2026");

  const [editingInfoId, setEditingInfoId] =
    useState(null);

  const [annualClientId, setAnnualClientId] =
    useState("");

  const [newAnnualTask, setNewAnnualTask] =
    useState("");

  const [newAnnualGroup, setNewAnnualGroup] =
    useState("Papildu darbi");

  const [newReportName, setNewReportName] =
    useState("");

  const [newReportPeriod, setNewReportPeriod] =
    useState("monthly");

  const [newReportDueDay, setNewReportDueDay] =
    useState("20");

  useEffect(() => {
    localStorage.setItem(
      STORAGE,
      JSON.stringify(data)
    );
  }, [data]);

  useEffect(() => {
    if (!running) return;

    const interval = setInterval(() => {
      setNow(Date.now());
    }, 1000);

    return () => clearInterval(interval);
  }, [running]);

  useEffect(() => {
    if (
      !clientId &&
      data.clients.length > 0
    ) {
      setClientId(data.clients[0].id);
    }

    if (
      !jobId &&
      data.jobs.length > 0
    ) {
      setJobId(data.jobs[0].id);
    }

    if (
      !annualClientId &&
      data.clients.length > 0
    ) {
      setAnnualClientId(
        data.clients[0].id
      );
    }
  }, [
    data.clients,
    data.jobs,
    clientId,
    jobId,
    annualClientId,
  ]);

  const selectedClient =
    data.clients.find(
      (c) => c.id === clientId
    );

  const selectedJob =
    data.jobs.find(
      (j) => j.id === jobId
    );

  const runningSeconds = running
    ? Math.max(
        0,
        Math.floor(
          (now - running.startedAt) /
            1000
        )
      )
    : 0;

  const runningTime =
    `${String(
      Math.floor(
        runningSeconds / 3600
      )
    ).padStart(2, "0")}:` +
    `${String(
      Math.floor(
        (runningSeconds % 3600) /
          60
      )
    ).padStart(2, "0")}:` +
    `${String(
      runningSeconds % 60
    ).padStart(2, "0")}`;

  const monthEntries = useMemo(
    () =>
      data.entries.filter(
        (entry) =>
          entry.date?.startsWith(
            selectedMonth
          )
      ),
    [data.entries, selectedMonth]
  );

  const totalMonthMinutes =
    monthEntries.reduce(
      (sum, entry) =>
        sum + entry.duration,
      0
    );

  function updateData(callback) {
    setData((current) =>
      callback(current)
    );
  }

  function addClient() {
    const name = newClient.trim();

    if (!name) {
      alert(
        "Ievadi klienta nosaukumu."
      );
      return;
    }

    const client = {
      id: uid("client"),
      name,
      vatPeriod: "monthly",
    };

    updateData((d) => ({
      ...d,
      clients: [
        ...d.clients,
        client,
      ],
    }));

    setClientId(client.id);

    setAnnualClientId(client.id);

    setNewClient("");
  }

  function deleteClient(id) {
    const client =
      data.clients.find(
        (c) => c.id === id
      );

    if (!client) return;

    const confirmed = window.confirm(
      `Vai tiešām dzēst klientu "${client.name}"?\n\n` +
        `Klients tiks izņemts no klientu saraksta, ` +
        `bet viņa iepriekšējie darba laika ieraksti ` +
        `NETIKS dzēsti.`
    );

    if (!confirmed) return;

    updateData((d) => {
      const statuses = {
        ...d.reportStatuses,
      };

      Object.keys(statuses).forEach(
        (key) => {
          if (
            key.includes(
              `|${id}|`
            )
          ) {
            delete statuses[key];
          }
        }
      );

      const annual = {
        ...d.annual,
      };

      delete annual[id];

      return {
        ...d,
        clients:
          d.clients.filter(
            (c) => c.id !== id
          ),
        reportStatuses:
          statuses,
        annual,
      };
    });

    if (clientId === id) {
      setClientId("");
    }

    if (annualClientId === id) {
      setAnnualClientId("");
    }
  }

  function addJob() {
    const name = newJob.trim();

    if (!name) {
      alert(
        "Ievadi darba nosaukumu."
      );
      return;
    }

    const job = {
      id: uid("job"),
      name,
    };

    updateData((d) => ({
      ...d,
      jobs: [
        ...d.jobs,
        job,
      ],
    }));

    setJobId(job.id);

    setNewJob("");
  }

  function deleteJob(id) {
    const job =
      data.jobs.find(
        (j) => j.id === id
      );

    if (!job) return;

    const confirmed = window.confirm(
      `Vai tiešām dzēst darbu "${job.name}"?\n\n` +
        `Esošie darba laika ieraksti NETIKS dzēsti.`
    );

    if (!confirmed) return;

    updateData((d) => ({
      ...d,
      jobs:
        d.jobs.filter(
          (j) => j.id !== id
        ),
    }));

    if (jobId === id) {
      setJobId("");
    }
  }

  function startWork() {
    if (!clientId) {
      alert(
        "Vispirms izvēlies klientu."
      );
      return;
    }

    if (!jobId) {
      alert(
        "Vispirms izvēlies darbu."
      );
      return;
    }

    if (running) {
      alert(
        "Darbs jau tiek uzskaitīts."
      );
      return;
    }

    setRunning({
      clientId,
      jobId,
      startedAt: Date.now(),
      date: todayISO(),
      note,
    });
  }

  function stopWork() {
    if (!running) return;

    const end = Date.now();

    const duration =
      Math.max(
        1,
        Math.round(
          (end -
            running.startedAt) /
            60000
        )
      );

    const entry = {
      id: uid("entry"),
      clientId:
        running.clientId,
      jobId:
        running.jobId,
      date:
        running.date,
      startedAt:
        new Date(
          running.startedAt
        ).toISOString(),
      endedAt:
        new Date(
          end
        ).toISOString(),
      duration,
      note:
        running.note || "",
    };

    updateData((d) => ({
      ...d,
      entries: [
        entry,
        ...d.entries,
      ],
    }));

    setRunning(null);

    setNote("");
  }

  function addManualEntry() {
    if (!clientId) {
      alert(
        "Vispirms izvēlies klientu."
      );
      return;
    }

    if (!jobId) {
      alert(
        "Vispirms izvēlies darbu."
      );
      return;
    }

    const minutes = Number(
      window.prompt(
        "Cik minūtes nostrādātas?",
        "60"
      )
    );

    if (
      !minutes ||
      minutes <= 0
    ) {
      return;
    }

    const date =
      window.prompt(
        "Datums YYYY-MM-DD",
        todayISO()
      );

    if (!date) return;

    const entry = {
      id: uid("entry"),
      clientId,
      jobId,
      date,
      startedAt: "",
      endedAt: "",
      duration:
        Math.round(minutes),
      note: note || "",
    };

    updateData((d) => ({
      ...d,
      entries: [
        entry,
        ...d.entries,
      ],
    }));

    setNote("");
  }

  function deleteEntry(id) {
    if (
      !window.confirm(
        "Dzēst šo darba laika ierakstu?"
      )
    ) {
      return;
    }

    updateData((d) => ({
      ...d,
      entries:
        d.entries.filter(
          (e) => e.id !== id
        ),
    }));
  }

  function exportExcel() {
    const dates =
      monthDates(
        selectedMonth
      );

    const rows =
      data.clients.map(
        (client) => {
          const row = {
            Klients:
              client.name,
          };

          dates.forEach(
            (date) => {
              const minutes =
                monthEntries
                  .filter(
                    (e) =>
                      e.clientId ===
                        client.id &&
                      e.date === date
                  )
                  .reduce(
                    (sum, e) =>
                      sum +
                      e.duration,
                    0
                  );

              row[date] = minutes
                ? Number(
                    minutesToHours(
                      minutes
                    ).toFixed(2)
                  )
                : "";
            }
          );

          const total =
            monthEntries
              .filter(
                (e) =>
                  e.clientId ===
                  client.id
              )
              .reduce(
                (sum, e) =>
                  sum + e.duration,
                0
              );

          row["Mēneša kopā (h)"] =
            roundHours(total);

          return row;
        }
      );

    const workbook =
      XLSX.utils.book_new();

    const worksheet =
      XLSX.utils.json_to_sheet(
        rows
      );

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Darba laiks"
    );

    const notes =
      monthEntries.map(
        (entry) => {
          const client =
            data.clients.find(
              (c) =>
                c.id ===
                entry.clientId
            );

          const job =
            data.jobs.find(
              (j) =>
                j.id ===
                entry.jobId
            );

          return {
            Datums:
              formatDate(
                entry.date
              ),
            Klients:
              client?.name ||
              "",
            "Paveiktais darbs":
              entry.note ||
              job?.name ||
              "",
          };
        }
      );

    const notesSheet =
      XLSX.utils.json_to_sheet(
        notes
      );

    XLSX.utils.book_append_sheet(
      workbook,
      notesSheet,
      "Piezīmes"
    );

    XLSX.writeFile(
      workbook,
      `gramatvedibas-darba-${selectedMonth}.xlsx`
    );
  }

  function saveInfo() {
    if (
      !infoCategory.trim() ||
      !infoTitle.trim() ||
      !infoValue.trim()
    ) {
      alert(
        "Aizpildi kategoriju, nosaukumu un informāciju."
      );
      return;
    }

    if (editingInfoId) {
      updateData((d) => ({
        ...d,
        info: d.info.map(
          (item) =>
            item.id ===
            editingInfoId
              ? {
                  ...item,
                  category:
                    infoCategory.trim(),
                  title:
                    infoTitle.trim(),
                  value:
                    infoValue.trim(),
                  year: infoYear,
                }
              : item
        ),
      }));
    } else {
      updateData((d) => ({
        ...d,
        info: [
          ...d.info,
          {
            id: uid("info"),
            category:
              infoCategory.trim(),
            title:
              infoTitle.trim(),
            value:
              infoValue.trim(),
            year: infoYear,
          },
        ],
      }));
    }

    resetInfoForm();
  }

  function resetInfoForm() {
    setInfoCategory("");
    setInfoTitle("");
    setInfoValue("");
    setInfoYear("2026");
    setEditingInfoId(null);
  }

  function editInfo(item) {
    setInfoCategory(
      item.category
    );

    setInfoTitle(
      item.title
    );

    setInfoValue(
      item.value
    );

    setInfoYear(
      item.year || ""
    );

    setEditingInfoId(
      item.id
    );

    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  }

  function deleteInfo(id) {
    if (
      !window.confirm(
        "Dzēst šo informāciju?"
      )
    ) {
      return;
    }

    updateData((d) => ({
      ...d,
      info:
        d.info.filter(
          (item) =>
            item.id !== id
        ),
    }));
  }

  function getAnnualTasks(
    clientIdValue
  ) {
    const custom =
      data.annual[
        clientIdValue
      ]?.custom || [];

    return [
      ...standardAnnualTasks.flatMap(
        (group) =>
          group.tasks.map(
            (text) => ({
              id:
                `${group.group}-${text}`,
              text,
              group:
                group.group,
              custom: false,
            })
          )
      ),
      ...custom,
    ];
  }

  function isAnnualDone(
    client,
    taskId
  ) {
    return Boolean(
      data.annual[client]
        ?.done?.[taskId]
    );
  }

  function toggleAnnualTask(
    client,
    taskId
  ) {
    updateData((d) => {
      const clientAnnual =
        d.annual[client] || {
          done: {},
          custom: [],
        };

      return {
        ...d,
        annual: {
          ...d.annual,
          [client]: {
            ...clientAnnual,
            done: {
              ...clientAnnual.done,
              [taskId]:
                !clientAnnual.done?.[
                  taskId
                ],
            },
          },
        },
      };
    });
  }

  function addAnnualTask() {
    if (
      !annualClientId ||
      !newAnnualTask.trim()
    ) {
      return;
    }

    const task = {
      id: uid("annual"),
      text:
        newAnnualTask.trim(),
      group:
        newAnnualGroup.trim() ||
        "Papildu darbi",
      custom: true,
    };

    updateData((d) => {
      const clientAnnual =
        d.annual[
          annualClientId
        ] || {
          done: {},
          custom: [],
        };

      return {
        ...d,
        annual: {
          ...d.annual,
          [annualClientId]: {
            ...clientAnnual,
            custom: [
              ...clientAnnual.custom,
              task,
            ],
          },
        },
      };
    });

    setNewAnnualTask("");
  }

  function deleteAnnualTask(
    taskId
  ) {
    updateData((d) => {
      const clientAnnual =
        d.annual[
          annualClientId
        ];

      if (!clientAnnual) {
        return d;
      }

      return {
        ...d,
        annual: {
          ...d.annual,
          [annualClientId]: {
            ...clientAnnual,
            custom:
              clientAnnual.custom.filter(
                (task) =>
                  task.id !==
                  taskId
              ),
          },
        },
      };
    });
  }

  function addReportType() {
    if (
      !newReportName.trim()
    ) {
      alert(
        "Ievadi atskaites nosaukumu."
      );
      return;
    }

    const dueDay =
      Number(
        newReportDueDay
      );

    if (
      !dueDay ||
      dueDay < 1 ||
      dueDay > 31
    ) {
      alert(
        "Termiņa dienai jābūt no 1 līdz 31."
      );
      return;
    }

    const report = {
      id: uid("report"),
      name:
        newReportName.trim(),
      period:
        newReportPeriod,
      dueDay,
      builtIn: false,
    };

    updateData((d) => ({
      ...d,
      reportTypes: [
        ...d.reportTypes,
        report,
      ],
    }));

    setNewReportName("");

    setNewReportPeriod(
      "monthly"
    );

    setNewReportDueDay("20");
  }

  function deleteReportType(
    reportId
  ) {
    const report =
      data.reportTypes.find(
        (r) =>
          r.id === reportId
      );

    if (
      !report ||
      report.builtIn
    ) {
      alert(
        "Iebūvēto atskaiti nevar dzēst."
      );
      return;
    }

    if (
      !window.confirm(
        `Dzēst atskaiti "${report.name}"?`
      )
    ) {
      return;
    }

    updateData((d) => ({
      ...d,
      reportTypes:
        d.reportTypes.filter(
          (r) =>
            r.id !== reportId
        ),
    }));
  }

  function reportStatusKey(
    reportId,
    clientIdValue
  ) {
    return `${selectedMonth}|${clientIdValue}|${reportId}`;
  }

  function getReportStatus(
    reportId,
    clientIdValue
  ) {
    return (
      data.reportStatuses[
        reportStatusKey(
          reportId,
          clientIdValue
        )
      ] || {
        submitted: false,
        taxSent: false,
      }
    );
  }

  function toggleReportStatus(
    reportId,
    clientIdValue,
    field
  ) {
    const key =
      reportStatusKey(
        reportId,
        clientIdValue
      );

    const current =
      getReportStatus(
        reportId,
        clientIdValue
      );

    updateData((d) => ({
      ...d,
      reportStatuses: {
        ...d.reportStatuses,
        [key]: {
          ...current,
          [field]:
            !current[field],
        },
      },
    }));
  }

  function periodLabel(
    period
  ) {
    if (
      period === "monthly"
    )
      return "Katru mēnesi";

    if (
      period === "quarterly"
    )
      return "Reizi ceturksnī";

    return "Reizi gadā";
  }

  function reportRelevant(
    client,
    report
  ) {
    if (
      report.id === "vat" &&
      client.vatPeriod ===
        "none"
    ) {
      return false;
    }

    return true;
  }

  function renderWork() {
    return (
      <>
        <div className="card">
          <div className="section-title">
            <div>
              <h2>Darba laiks</h2>

              <p>
                Uzskaiti laiku pēc
                klienta un veicamā
                darba.
              </p>
            </div>

            <Clock3 size={28} />
          </div>

          <div className="form-grid">
            <div>
              <label>
                Klients
              </label>

              <div className="inline">
                <select
                  value={clientId}
                  onChange={(e) =>
                    setClientId(
                      e.target.value
                    )
                  }
                >
                  <option value="">
                    Izvēlies klientu
                  </option>

                  {data.clients.map(
                    (client) => (
                      <option
                        key={
                          client.id
                        }
                        value={
                          client.id
                        }
                      >
                        {client.name}
                      </option>
                    )
                  )}
                </select>

                {clientId && (
                  <button
                    className="danger-btn"
                    onClick={() =>
                      deleteClient(
                        clientId
                      )
                    }
                    title="Dzēst izvēlēto klientu"
                  >
                    <Trash2
                      size={17}
                    />
                    Dzēst
                  </button>
                )}
              </div>

              <div
                className="inline"
                style={{
                  marginTop: 10,
                }}
              >
                <input
                  type="text"
                  placeholder="Jauns klients..."
                  value={
                    newClient
                  }
                  onChange={(e) =>
                    setNewClient(
                      e.target.value
                    )
                  }
                />

                <button
                  className="icon-btn"
                  onClick={
                    addClient
                  }
                  title="Pievienot klientu"
                >
                  <Plus size={22} />
                </button>
              </div>
            </div>

            <div>
              <label>
                Darbs
              </label>

              <select
                value={jobId}
                onChange={(e) =>
                  setJobId(
                    e.target.value
                  )
                }
              >
                <option value="">
                  Izvēlies darbu
                </option>

                {data.jobs.map(
                  (job) => (
                    <option
                      key={job.id}
                      value={job.id}
                    >
                      {job.name}
                    </option>
                  )
                )}
              </select>

              <div
                className="inline"
                style={{
                  marginTop: 10,
                }}
              >
                <input
                  type="text"
                  placeholder="Jauns darbs..."
                  value={newJob}
                  onChange={(e) =>
                    setNewJob(
                      e.target.value
                    )
                  }
                />

                <button
                  className="icon-btn"
                  onClick={addJob}
                  title="Pievienot darbu"
                >
                  <Plus size={22} />
                </button>
              </div>
            </div>
          </div>

          <div
            style={{
              marginTop: 24,
            }}
          >
            <label>
              Ko izdarīji?
            </label>

            <textarea
              value={note}
              onChange={(e) =>
                setNote(
                  e.target.value
                )
              }
              placeholder="Piemēram: pārbaudīti bankas izraksti, sagatavota PVN deklarācija..."
            />
          </div>

          {!running ? (
            <button
              className="primary-btn"
              onClick={
                startWork
              }
            >
              <Clock3
                size={21}
              />
              Sākt darbu
            </button>
          ) : (
            <div className="timer-box">
              <div>
                <strong>
                  {selectedClient?.name ||
                    "Klients"}
                </strong>

                <span>
                  {selectedJob?.name ||
                    "Darbs"}
                </span>
              </div>

              <div className="timer">
                {runningTime}
              </div>

              <div className="timer-actions">
                <button
                  className="primary-btn"
                  onClick={
                    stopWork
                  }
                >
                  <CheckCircle2
                    size={21}
                  />
                  Pabeigt darbu
                </button>
              </div>
            </div>
          )}

          <div
            style={{
              marginTop: 12,
            }}
          >
            <button
              className="secondary-btn"
              onClick={
                addManualEntry
              }
            >
              <Plus size={18} />
              Pievienot laiku
              manuāli
            </button>
          </div>
        </div>

        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Klientu pārvaldība
              </h2>

              <p>
                Šeit vari jebkurā
                laikā dzēst klientu,
                ar kuru vairs
                nestrādā.
              </p>
            </div>

            <Users size={27} />
          </div>

          {data.clients.length ===
          0 ? (
            <div className="empty-state">
              <Users size={35} />

              <strong>
                Nav pievienotu
                klientu
              </strong>

              <span>
                Pievieno pirmo klientu
                augstāk.
              </span>
            </div>
          ) : (
            <div className="manage-list">
              {data.clients.map(
                (client) => (
                  <div
                    className="manage-row"
                    key={
                      client.id
                    }
                  >
                    <div>
                      <strong>
                        {
                          client.name
                        }
                      </strong>

                      <span>
                        PVN:{" "}
                        {client.vatPeriod ===
                        "monthly"
                          ? "katru mēnesi"
                          : client.vatPeriod ===
                            "quarterly"
                          ? "ceturksnī"
                          : "nav PVN"}
                      </span>
                    </div>

                    <button
                      className="danger-btn"
                      onClick={() =>
                        deleteClient(
                          client.id
                        )
                      }
                    >
                      <Trash2
                        size={17}
                      />
                      Dzēst
                    </button>
                  </div>
                )
              )}
            </div>
          )}
        </div>

        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Darbu pārvaldība
              </h2>

              <p>
                Dzēšot darbu,
                vēsturiskie ieraksti
                saglabājas.
              </p>
            </div>

            <BriefcaseBusiness
              size={27}
            />
          </div>

          <div className="manage-list">
            {data.jobs.map(
              (job) => (
                <div
                  className="manage-row"
                  key={job.id}
                >
                  <div>
                    <strong>
                      {job.name}
                    </strong>
                  </div>

                  <button
                    className="danger-btn"
                    onClick={() =>
                      deleteJob(
                        job.id
                      )
                    }
                  >
                    <Trash2
                      size={17}
                    />
                    Dzēst
                  </button>
                </div>
              )
            )}
          </div>
        </div>

        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Pēdējie ieraksti
              </h2>

              <p>
                Saglabātā darba laika
                uzskaite.
              </p>
            </div>
          </div>

          <div className="entries">
            {data.entries
              .slice(0, 15)
              .map((entry) => {
                const client =
                  data.clients.find(
                    (c) =>
                      c.id ===
                      entry.clientId
                  );

                const job =
                  data.jobs.find(
                    (j) =>
                      j.id ===
                      entry.jobId
                  );

                return (
                  <div
                    className="entry"
                    key={
                      entry.id
                    }
                  >
                    <div>
                      <strong>
                        {client?.name ||
                          "Dzēsts klients"}
                      </strong>

                      <span>
                        {job?.name ||
                          "Dzēsts darbs"}
                      </span>

                      <small>
                        {formatDate(
                          entry.date
                        )}{" "}
                        ·{" "}
                        {formatDuration(
                          entry.duration
                        )}
                      </small>

                      {entry.note && (
                        <small>
                          {
                            entry.note
                          }
                        </small>
                      )}
                    </div>

                    <button
                      className="danger-btn"
                      onClick={() =>
                        deleteEntry(
                          entry.id
                        )
                      }
                    >
                      <Trash2
                        size={16}
                      />
                    </button>
                  </div>
                );
              })}

            {!data.entries
              .length && (
              <p className="muted">
                Vēl nav saglabātu
                darba laika ierakstu.
              </p>
            )}
          </div>
        </div>
      </>
    );
  }

  function renderTable() {
    const dates =
      monthDates(
        selectedMonth
      );

    return (
      <div className="card">
        <div className="section-title">
          <div>
            <h2>
              Darba laika atskaite
            </h2>

            <p>
              {monthLabel(
                selectedMonth
              )}{" "}
              · mēneša kopsummas
              noapaļotas līdz pilnām
              stundām.
            </p>
          </div>

          <button
            className="secondary-btn"
            onClick={
              exportExcel
            }
          >
            <Download
              size={18}
            />
            Excel
          </button>
        </div>

        <div
          className="inline"
          style={{
            marginBottom: 20,
          }}
        >
          <input
            type="month"
            value={
              selectedMonth
            }
            onChange={(e) =>
              setSelectedMonth(
                e.target.value
              )
            }
          />
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>
                  Klients
                </th>

                {dates.map(
                  (date) => (
                    <th
                      key={date}
                    >
                      {Number(
                        date.slice(
                          -2
                        )
                      )}
                    </th>
                  )
                )}

                <th>
                  Mēnesī
                </th>
              </tr>
            </thead>

            <tbody>
              {data.clients.map(
                (client) => {
                  const clientMinutes =
                    monthEntries
                      .filter(
                        (e) =>
                          e.clientId ===
                          client.id
                      )
                      .reduce(
                        (sum, e) =>
                          sum +
                          e.duration,
                        0
                      );

                  return (
                    <tr
                      key={
                        client.id
                      }
                    >
                      <td>
                        <strong>
                          {
                            client.name
                          }
                        </strong>
                      </td>

                      {dates.map(
                        (date) => {
                          const minutes =
                            monthEntries
                              .filter(
                                (e) =>
                                  e.clientId ===
                                    client.id &&
                                  e.date ===
                                    date
                              )
                              .reduce(
                                (
                                  sum,
                                  e
                                ) =>
                                  sum +
                                  e.duration,
                                0
                              );

                          return (
                            <td
                              key={
                                date
                              }
                            >
                              {minutes
                                ? minutesToHours(
                                    minutes
                                  ).toFixed(
                                    2
                                  )
                                : ""}
                            </td>
                          );
                        }
                      )}

                      <td>
                        <strong>
                          {
                            roundHours(
                              clientMinutes
                            )
                          }{" "}
                          h
                        </strong>
                      </td>
                    </tr>
                  );
                }
              )}

              <tr>
                <td>
                  <strong>
                    KOPĀ
                  </strong>
                </td>

                {dates.map(
                  (date) => {
                    const minutes =
                      monthEntries
                        .filter(
                          (e) =>
                            e.date ===
                            date
                        )
                        .reduce(
                          (sum, e) =>
                            sum +
                            e.duration,
                          0
                        );

                    return (
                      <td
                        key={
                          date
                        }
                      >
                        {minutes
                          ? minutesToHours(
                              minutes
                            ).toFixed(
                              2
                            )
                          : ""}
                      </td>
                    );
                  }
                )}

                <td>
                  <strong>
                    {
                      roundHours(
                        totalMonthMinutes
                      )
                    }{" "}
                    h
                  </strong>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="notes-box">
          <h3>
            Paveiktais darbs
          </h3>

          {monthEntries.length ===
          0 ? (
            <p className="muted">
              Šajā mēnesī vēl nav
              ierakstu.
            </p>
          ) : (
            monthEntries
              .slice()
              .sort(
                (a, b) =>
                  a.date.localeCompare(
                    b.date
                  )
              )
              .map((entry) => {
                const client =
                  data.clients.find(
                    (c) =>
                      c.id ===
                      entry.clientId
                  );

                const job =
                  data.jobs.find(
                    (j) =>
                      j.id ===
                      entry.jobId
                  );

                return (
                  <div
                    className="note-row"
                    key={
                      entry.id
                    }
                  >
                    <strong>
                      {formatDate(
                        entry.date
                      )}{" "}
                      ·{" "}
                      {
                        client?.name
                      }
                    </strong>

                    <span>
                      {entry.note ||
                        job?.name ||
                        ""}
                    </span>
                  </div>
                );
              })
          )}
        </div>
      </div>
    );
  }

  function renderCalendar() {
    const [
      year,
      month,
    ] = selectedMonth
      .split("-")
      .map(Number);

    const firstDay =
      new Date(
        year,
        month - 1,
        1
      );

    const daysInMonth =
      new Date(
        year,
        month,
        0
      ).getDate();

    const startDay =
      (firstDay.getDay() +
        6) %
      7;

    const cells = [];

    for (
      let i = 0;
      i < startDay;
      i++
    ) {
      cells.push(
        <div
          className="calendar-day empty"
          key={`empty-${i}`}
        />
      );
    }

    for (
      let day = 1;
      day <= daysInMonth;
      day++
    ) {
      const date =
        new Date(
          year,
          month - 1,
          day
        );

      const weekday =
        date.getDay();

      const weekend =
        weekday === 0 ||
        weekday === 6;

      cells.push(
        <div
          className={`calendar-day ${
            weekend ? "red" : ""
          }`}
          key={day}
        >
          {day}
        </div>
      );
    }

    return (
      <div className="card">
        <div className="section-title">
          <div>
            <h2>
              Kalendārs
            </h2>

            <p>
              Darba dienas un
              nedēļas nogales.
            </p>
          </div>

          <CalendarDays
            size={28}
          />
        </div>

        <div
          className="inline"
          style={{
            marginBottom: 22,
          }}
        >
          <button
            className="icon-btn"
            onClick={() => {
              const [
                y,
                m,
              ] =
                selectedMonth
                  .split("-")
                  .map(Number);

              const d =
                new Date(
                  y,
                  m - 2,
                  1
                );

              setSelectedMonth(
                `${d.getFullYear()}-${String(
                  d.getMonth() + 1
                ).padStart(
                  2,
                  "0"
                )}`
              );
            }}
          >
            <ChevronLeft
              size={21}
            />
          </button>

          <input
            type="month"
            value={
              selectedMonth
            }
            onChange={(e) =>
              setSelectedMonth(
                e.target.value
              )
            }
          />

          <button
            className="icon-btn"
            onClick={() => {
              const [
                y,
                m,
              ] =
                selectedMonth
                  .split("-")
                  .map(Number);

              const d =
                new Date(
                  y,
                  m,
                  1
                );

              setSelectedMonth(
                `${d.getFullYear()}-${String(
                  d.getMonth() + 1
                ).padStart(
                  2,
                  "0"
                )}`
              );
            }}
          >
            <ChevronRight
              size={21}
            />
          </button>
        </div>

        <div className="calendar">
          {[
            "P",
            "O",
            "T",
            "C",
            "P",
            "S",
            "Sv",
          ].map(
            (day) => (
              <div
                className="calendar-head"
                key={day}
              >
                {day}
              </div>
            )
          )}

          {cells}
        </div>

        <div className="legend">
          <span className="legend-red" />
          Nedēļas nogale
        </div>
      </div>
    );
  }

  function renderInfo() {
    const categories = [
      ...new Set(
        data.info.map(
          (item) =>
            item.category
        )
      ),
    ];

    return (
      <>
        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Informācija
              </h2>

              <p>
                Tava grāmatvedības
                informācijas datubāze.
                Informāciju vari
                papildināt un labot.
              </p>
            </div>

            <Info size={28} />
          </div>

          <div className="info-form">
            <div className="form-grid">
              <div>
                <label>
                  Kategorija
                </label>

                <input
                  type="text"
                  placeholder="Piemēram, PVN"
                  value={
                    infoCategory
                  }
                  onChange={(e) =>
                    setInfoCategory(
                      e.target.value
                    )
                  }
                />
              </div>

              <div>
                <label>
                  Nosaukums
                </label>

                <input
                  type="text"
                  placeholder="Piemēram, PVN standarta likme"
                  value={
                    infoTitle
                  }
                  onChange={(e) =>
                    setInfoTitle(
                      e.target.value
                    )
                  }
                />
              </div>
            </div>

            <div
              style={{
                marginTop: 16,
              }}
            >
              <label>
                Informācija
              </label>

              <textarea
                value={
                  infoValue
                }
                onChange={(e) =>
                  setInfoValue(
                    e.target.value
                  )
                }
                placeholder="Ievadi informāciju..."
              />
            </div>

            <div className="form-grid">
              <div>
                <label>
                  Gads
                </label>

                <input
                  type="text"
                  value={
                    infoYear
                  }
                  onChange={(e) =>
                    setInfoYear(
                      e.target.value
                    )
                  }
                />
              </div>

              <div
                style={{
                  display:
                    "flex",
                  alignItems:
                    "flex-end",
                  gap: 10,
                }}
              >
                <button
                  className="primary-btn"
                  onClick={
                    saveInfo
                  }
                >
                  {editingInfoId ? (
                    <>
                      <Save
                        size={19}
                      />
                      Saglabāt
                      izmaiņas
                    </>
                  ) : (
                    <>
                      <Plus
                        size={19}
                      />
                      Pievienot
                      informāciju
                    </>
                  )}
                </button>

                {editingInfoId && (
                  <button
                    className="secondary-btn"
                    onClick={
                      resetInfoForm
                    }
                  >
                    <X
                      size={18}
                    />
                    Atcelt
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {categories.map(
          (category) => (
            <div
              className="card"
              key={category}
            >
              <div className="section-title">
                <div>
                  <h2>
                    {category}
                  </h2>
                </div>

                <FileText
                  size={25}
                />
              </div>

              <div className="info-grid">
                {data.info
                  .filter(
                    (item) =>
                      item.category ===
                      category
                  )
                  .map(
                    (item) => (
                      <div
                        className="info-card"
                        key={
                          item.id
                        }
                      >
                        <div
                          style={{
                            display:
                              "flex",
                            justifyContent:
                              "space-between",
                            gap: 10,
                          }}
                        >
                          <strong>
                            {
                              item.title
                            }
                          </strong>

                          {item.year && (
                            <span>
                              {
                                item.year
                              }
                            </span>
                          )}
                        </div>

                        <span>
                          {
                            item.value
                          }
                        </span>

                        <div
                          style={{
                            display:
                              "flex",
                            gap: 7,
                            marginTop:
                              "auto",
                          }}
                        >
                          <button
                            className="secondary-btn"
                            onClick={() =>
                              editInfo(
                                item
                              )
                            }
                          >
                            Labot
                          </button>

                          <button
                            className="danger-btn"
                            onClick={() =>
                              deleteInfo(
                                item.id
                              )
                            }
                          >
                            <Trash2
                              size={
                                15
                              }
                            />
                          </button>
                        </div>
                      </div>
                    )
                  )}
              </div>
            </div>
          )
        )}
      </>
    );
  }

  function renderAnnual() {
    if (
      data.clients.length ===
      0
    ) {
      return (
        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Gada pārskats
              </h2>

              <p>
                Vispirms pievieno
                klientu.
              </p>
            </div>

            <ClipboardCheck
              size={28}
            />
          </div>

          <p className="muted">
            Gada pārskata
            kontrolsaraksts tiks
            izveidots katram
            pievienotajam klientam.
          </p>
        </div>
      );
    }

    const tasks =
      getAnnualTasks(
        annualClientId
      );

    const completed =
      tasks.filter(
        (task) =>
          isAnnualDone(
            annualClientId,
            task.id
          )
      ).length;

    const progress =
      tasks.length
        ? Math.round(
            (completed /
              tasks.length) *
              100
          )
        : 0;

    const groups = [
      ...new Set(
        tasks.map(
          (task) =>
            task.group
        )
      ),
    ];

    return (
      <div className="card">
        <div className="section-title">
          <div>
            <h2>
              Gada pārskats
            </h2>

            <p>
              Profesionāls
              kontrolsaraksts katram
              klientam.
            </p>
          </div>

          <ClipboardCheck
            size={28}
          />
        </div>

        <label>
          Klients
        </label>

        <select
          value={
            annualClientId
          }
          onChange={(e) =>
            setAnnualClientId(
              e.target.value
            )
          }
        >
          {data.clients.map(
            (client) => (
              <option
                key={
                  client.id
                }
                value={
                  client.id
                }
              >
                {client.name}
              </option>
            )
          )}
        </select>

        <div
          style={{
            marginTop: 22,
            padding: 20,
            borderRadius: 20,
            background:
              "#edf8f1",
          }}
        >
          <div
            style={{
              display:
                "flex",
              justifyContent:
                "space-between",
              marginBottom: 10,
            }}
          >
            <strong>
              Izpildīts:{" "}
              {completed} /{" "}
              {tasks.length}
            </strong>

            <strong>
              {progress}%
            </strong>
          </div>

          <div
            style={{
              height: 10,
              background:
                "#d7e8dc",
              borderRadius: 99,
              overflow:
                "hidden",
            }}
          >
            <div
              style={{
                width: `${progress}%`,
                height: "100%",
                background:
                  "#55ad79",
                borderRadius: 99,
                transition:
                  "width .2s",
              }}
            />
          </div>
        </div>

        <div className="annual-list">
          {groups.map(
            (group) => (
              <div
                className="annual-group"
                key={group}
              >
                <h3>
                  {group}
                </h3>

                <div className="checklist">
                  {tasks
                    .filter(
                      (task) =>
                        task.group ===
                        group
                    )
                    .map(
                      (task) => (
                        <label
                          className="check-row"
                          key={
                            task.id
                          }
                        >
                          <input
                            type="checkbox"
                            checked={isAnnualDone(
                              annualClientId,
                              task.id
                            )}
                            onChange={() =>
                              toggleAnnualTask(
                                annualClientId,
                                task.id
                              )
                            }
                          />

                          <span>
                            {
                              task.text
                            }
                          </span>

                          {isAnnualDone(
                            annualClientId,
                            task.id
                          ) && (
                            <Check
                              size={
                                19
                              }
                            />
                          )}

                          {task.custom && (
                            <button
                              type="button"
                              className="danger-btn"
                              onClick={(
                                e
                              ) => {
                                e.preventDefault();

                                deleteAnnualTask(
                                  task.id
                                );
                              }}
                            >
                              <Trash2
                                size={
                                  15
                                }
                              />
                            </button>
                          )}
                        </label>
                      )
                    )}
                </div>
              </div>
            )
          )}
        </div>

        <div className="notes-box">
          <h3>
            Pievienot klientam
            savu uzdevumu
          </h3>

          <div className="form-grid">
            <div>
              <label>
                Grupa
              </label>

              <input
                type="text"
                value={
                  newAnnualGroup
                }
                onChange={(e) =>
                  setNewAnnualGroup(
                    e.target.value
                  )
                }
              />
            </div>

            <div>
              <label>
                Uzdevums
              </label>

              <input
                type="text"
                placeholder="Piemēram, pārbaudīt aizdevuma līgumu"
                value={
                  newAnnualTask
                }
                onChange={(e) =>
                  setNewAnnualTask(
                    e.target.value
                  )
                }
              />
            </div>
          </div>

          <button
            className="secondary-btn"
            style={{
              marginTop: 15,
            }}
            onClick={
              addAnnualTask
            }
          >
            <Plus size={18} />
            Pievienot uzdevumu
          </button>
        </div>
      </div>
    );
  }

  function renderDeadlines() {
    if (
      data.clients.length ===
      0
    ) {
      return (
        <>
          <div className="card">
            <div className="section-title">
              <div>
                <h2>
                  Termiņi un
                  atskaites
                </h2>

                <p>
                  Pievieno jebkuru
                  atskaiti, kuru
                  vēlies pārvaldīt.
                </p>
              </div>

              <FileText
                size={28}
              />
            </div>

            <p className="muted">
              Vispirms pievieno
              vismaz vienu klientu
              sadaļā "Darba laiks".
            </p>
          </div>

          <div className="card">
            <ReportTypeForm
              newReportName={
                newReportName
              }
              setNewReportName={
                setNewReportName
              }
              newReportPeriod={
                newReportPeriod
              }
              setNewReportPeriod={
                setNewReportPeriod
              }
              newReportDueDay={
                newReportDueDay
              }
              setNewReportDueDay={
                setNewReportDueDay
              }
              addReportType={
                addReportType
              }
            />
          </div>
        </>
      );
    }

    return (
      <>
        <div className="card">
          <ReportTypeForm
            newReportName={
              newReportName
            }
            setNewReportName={
              setNewReportName
            }
            newReportPeriod={
              newReportPeriod
            }
            setNewReportPeriod={
              setNewReportPeriod
            }
            newReportDueDay={
              newReportDueDay
            }
            setNewReportDueDay={
              setNewReportDueDay
            }
            addReportType={
              addReportType
            }
          />
        </div>

        <div className="card">
          <div className="section-title">
            <div>
              <h2>
                Atskaites
              </h2>

              <p>
                Visas atskaites,
                kuras esi pievienojusi.
              </p>
            </div>

            <Settings2
              size={27}
            />
          </div>

          <div className="report-list">
            {data.reportTypes.map(
              (report) => (
                <div
                  className="report-card"
                  key={
                    report.id
                  }
                >
                  <div className="report-card-header">
                    <div>
                      <strong>
                        {
                          report.name
                        }
                      </strong>

                      <span>
                        {periodLabel(
                          report.period
                        )}{" "}
                        · līdz{" "}
                        {
                          report.dueDay
                        }
                        . datumam
                      </span>
                    </div>

                    {!report.builtIn && (
                      <button
                        className="danger-btn"
                        onClick={() =>
                          deleteReportType(
                            report.id
                          )
                        }
                      >
                        <Trash2
                          size={
                            16
                          }
                        />
                        Dzēst
                      </button>
                    )}
                  </div>

                  <div className="report-client-list">
                    {data.clients.map(
                      (client) => {
                        if (
                          !reportRelevant(
                            client,
                            report
                          )
                        ) {
                          return null;
                        }

                        const status =
                          getReportStatus(
                            report.id,
                            client.id
                          );

                        return (
                          <div
                            className="report-client-row"
                            key={
                              client.id
                            }
                          >
                            <strong>
                              {
                                client.name
                              }
                            </strong>

                            <label>
                              <input
                                type="checkbox"
                                checked={
                                  status.submitted
                                }
                                onChange={() =>
                                  toggleReportStatus(
                                    report.id,
                                    client.id,
                                    "submitted"
                                  )
                                }
                              />

                              Iesniegta
                            </label>

                            <label>
                              <input
                                type="checkbox"
                                checked={
                                  status.taxSent
                                }
                                onChange={() =>
                                  toggleReportStatus(
                                    report.id,
                                    client.id,
                                    "taxSent"
                                  )
                                }
                              />

                              Nodokļa informācija
                            </label>

                            {status.submitted && (
                              <CheckCircle2
                                size={
                                  20
                                }
                              />
                            )}
                          </div>
                        );
                      }
                    )}
                  </div>
                </div>
              )
            )}
          </div>
        </div>
      </>
    );
  }

  function renderPage() {
    if (
      page === "work"
    )
      return renderWork();

    if (
      page === "table"
    )
      return renderTable();

    if (
      page === "calendar"
    )
      return renderCalendar();

    if (
      page === "info"
    )
      return renderInfo();

    if (
      page === "annual"
    )
      return renderAnnual();

    if (
      page === "deadlines"
    )
      return renderDeadlines();

    return null;
  }

  const navItems = [
    {
      id: "work",
      label: "Darba laiks",
      icon: Clock3,
    },
    {
      id: "table",
      label: "Atskaite",
      icon: ClipboardCheck,
    },
    {
      id: "calendar",
      label: "Kalendārs",
      icon: CalendarDays,
    },
    {
      id: "info",
      label: "Informācija",
      icon: Info,
    },
    {
      id: "annual",
      label: "Gada pārskats",
      icon: FileText,
    },
    {
      id: "deadlines",
      label: "Termiņi",
      icon: CheckCircle2,
    },
  ];

  return (
    <div className="app">
      <header className="topbar">
        <h1>
          Grāmatvedības darba
          pārvaldība
        </h1>

        <p>
          Darba laiks, klienti,
          atskaites, termiņi un
          grāmatvedības informācija
          vienuviet.
        </p>
      </header>

      <nav className="nav">
        {navItems.map(
          (item) => {
            const Icon =
              item.icon;

            return (
              <button
                key={
                  item.id
                }
                className={`nav-btn ${
                  page ===
                  item.id
                    ? "active"
                    : ""
                }`}
                onClick={() =>
                  setPage(
                    item.id
                  )
                }
              >
                <Icon />

                <span>
                  {
                    item.label
                  }
                </span>
              </button>
            );
          }
        )}
      </nav>

      <main className="content">
        {renderPage()}
      </main>
    </div>
  );
}

function ReportTypeForm({
  newReportName,
  setNewReportName,
  newReportPeriod,
  setNewReportPeriod,
  newReportDueDay,
  setNewReportDueDay,
  addReportType,
}) {
  return (
    <>
      <div className="section-title">
        <div>
          <h2>
            Pievienot atskaiti
          </h2>

          <p>
            Vari izveidot arī savu
            atskaiti, kas nav
            iebūvēta sistēmā.
          </p>
        </div>

        <FileText size={28} />
      </div>

      <div className="form-grid">
        <div>
          <label>
            Atskaites nosaukums
          </label>

          <input
            type="text"
            placeholder="Piemēram, Statistikas pārskats"
            value={
              newReportName
            }
            onChange={(e) =>
              setNewReportName(
                e.target.value
              )
            }
          />
        </div>

        <div>
          <label>
            Periods
          </label>

          <select
            value={
              newReportPeriod
            }
            onChange={(e) =>
              setNewReportPeriod(
                e.target.value
              )
            }
          >
            <option value="monthly">
              Katru mēnesi
            </option>

            <option value="quarterly">
              Reizi ceturksnī
            </option>

            <option value="yearly">
              Reizi gadā
            </option>
          </select>
        </div>
      </div>

      <div
        className="form-grid"
        style={{
          marginTop: 16,
        }}
      >
        <div>
          <label>
            Termiņa diena
          </label>

          <input
            type="text"
            inputMode="numeric"
            value={
              newReportDueDay
            }
            onChange={(e) =>
              setNewReportDueDay(
                e.target.value
              )
            }
          />
        </div>

        <div
          style={{
            display: "flex",
            alignItems:
              "flex-end",
          }}
        >
          <button
            className="primary-btn"
            onClick={
              addReportType
            }
          >
            <Plus size={19} />
            Pievienot atskaiti
          </button>
        </div>
      </div>
    </>
  );
}

createRoot(
  document.getElementById(
    "root"
  )
).render(<App />);
