import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Clock3,
  Table2,
  CalendarDays,
  BookOpen,
  ClipboardCheck,
  Bell,
  Plus,
  Play,
  Square,
  Pause,
  Download,
  Check,
  Trash2,
  Save,
} from "lucide-react";
import "./styles.css";

const STORAGE = "gramatvedibas-darba-v2";

const defaultData = {
  clients: [
    { id: "c1", name: "SIA ABC", vat: "monthly" },
    { id: "c2", name: "SIA XYZ", vat: "quarterly" },
    { id: "c3", name: "IK Anna", vat: "none" },
  ],
  jobs: [
    "Grāmatojumi",
    "PVN deklarācija",
    "Algu aprēķins",
    "Bankas kontrole",
    "Salīdzināšanas akti",
  ],
  entries: [],
  reports: {},
  annual: {},
  notes: {},
};

function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE);
    return saved ? JSON.parse(saved) : defaultData;
  } catch {
    return defaultData;
  }
}

function App() {
  const [data, setData] = useState(loadData);
  const [page, setPage] = useState("work");
  const [clientId, setClientId] = useState("");
  const [job, setJob] = useState("");
  const [note, setNote] = useState("");
  const [timer, setTimer] = useState(null);
  const [now, setNow] = useState(Date.now());

  useEffect(() => {
    localStorage.setItem(STORAGE, JSON.stringify(data));
  }, [data]);

  useEffect(() => {
    const id = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, []);

  const client = data.clients.find((c) => c.id === clientId);

  function addClient() {
    const name = prompt("Klienta nosaukums:");
    if (!name?.trim()) return;

    const vat = prompt(
      "PVN periods: monthly = katru mēnesi, quarterly = reizi ceturksnī, none = nav PVN"
    );

    const newClient = {
      id: Date.now().toString(),
      name: name.trim(),
      vat:
        vat === "quarterly"
          ? "quarterly"
          : vat === "none"
          ? "none"
          : "monthly",
    };

    setData((d) => ({
      ...d,
      clients: [...d.clients, newClient],
    }));

    setClientId(newClient.id);
  }

  function addJob() {
    const name = prompt("Darba veids:");
    if (!name?.trim()) return;

    setData((d) => ({
      ...d,
      jobs: [...d.jobs, name.trim()],
    }));

    setJob(name.trim());
  }

  function startTimer() {
    if (!clientId || !job) {
      alert("Vispirms izvēlies klientu un darbu.");
      return;
    }

    setTimer({
      start: Date.now(),
      clientId,
      job,
    });
  }

  function stopTimer() {
    if (!timer) return;

    const end = Date.now();
    const duration = Math.max(0, end - timer.start);

    const entry = {
      id: Date.now().toString(),
      date: new Date(timer.start).toISOString().slice(0, 10),
      clientId: timer.clientId,
      job: timer.job,
      start: new Date(timer.start).toLocaleTimeString("lv-LV", {
        hour: "2-digit",
        minute: "2-digit",
      }),
      end: new Date(end).toLocaleTimeString("lv-LV", {
        hour: "2-digit",
        minute: "2-digit",
      }),
      duration,
      note,
    };

    setData((d) => ({
      ...d,
      entries: [...d.entries, entry],
    }));

    setTimer(null);
    setNote("");
  }

  function pauseTimer() {
    alert("Pauzes funkciju pievienosim nākamajā versijā.");
  }

  function formatDuration(ms) {
    const totalSeconds = Math.floor(ms / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
      2,
      "0"
    )}:${String(seconds).padStart(2, "0")}`;
  }

  function exportExcel() {
    const rows = data.entries.map((e) => ({
      Datums: e.date,
      Klients:
        data.clients.find((c) => c.id === e.clientId)?.name || "Nezināms",
      Darbs: e.job,
      Sākums: e.start,
      Beigas: e.end,
      Stundas: (e.duration / 3600000).toFixed(2),
      Piezīme: e.note || "",
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Darba laiks");
    XLSX.writeFile(wb, "darba-laiks.xlsx");
  }

  const nav = [
    ["work", "Darba laiks", Clock3],
    ["table", "Atskaite", Table2],
    ["calendar", "Kalendārs", CalendarDays],
    ["info", "Informācija", BookOpen],
    ["annual", "Gada pārskats", ClipboardCheck],
    ["deadlines", "Termiņi", Bell],
  ];

  return (
    <div className="app">
      <header className="topbar">
        <div>
          <h1>Grāmatvedības darba pārvaldība</h1>
          <p>Darba laiks, klienti, atskaites un termiņi</p>
        </div>
      </header>

      <nav className="nav">
        {nav.map(([id, label, Icon]) => (
          <button
            key={id}
            className={page === id ? "nav-btn active" : "nav-btn"}
            onClick={() => setPage(id)}
          >
            <Icon size={18} />
            <span>{label}</span>
          </button>
        ))}
      </nav>

      <main className="content">
        {page === "work" && (
          <WorkPage
            data={data}
            clientId={clientId}
            setClientId={setClientId}
            job={job}
            setJob={setJob}
            note={note}
            setNote={setNote}
            timer={timer}
            now={now}
            addClient={addClient}
            addJob={addJob}
            startTimer={startTimer}
            stopTimer={stopTimer}
            pauseTimer={pauseTimer}
            formatDuration={formatDuration}
          />
        )}

        {page === "table" && (
          <TablePage data={data} exportExcel={exportExcel} />
        )}

        {page === "calendar" && <CalendarPage />}

        {page === "info" && <InfoPage />}

        {page === "annual" && <AnnualPage data={data} setData={setData} />}

        {page === "deadlines" && (
          <DeadlinesPage data={data} setData={setData} />
        )}
      </main>
    </div>
  );
}

function WorkPage({
  data,
  clientId,
  setClientId,
  job,
  setJob,
  note,
  setNote,
  timer,
  now,
  addClient,
  addJob,
  startTimer,
  stopTimer,
  pauseTimer,
  formatDuration,
}) {
  return (
    <>
      <section className="card">
        <div className="section-title">
          <div>
            <h2>Darba laiks</h2>
            <p>Izvēlies klientu un darbu un sāc uzskaiti.</p>
          </div>
        </div>

        <div className="form-grid">
          <div>
            <label>Klients</label>
            <div className="inline">
              <select
                value={clientId}
                onChange={(e) => setClientId(e.target.value)}
              >
                <option value="">Izvēlies klientu</option>
                {data.clients.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name}
                  </option>
                ))}
              </select>

              <button className="icon-btn" onClick={addClient}>
                <Plus size={18} />
              </button>
            </div>
          </div>

          <div>
            <label>Darbs</label>
            <div className="inline">
              <select value={job} onChange={(e) => setJob(e.target.value)}>
                <option value="">Izvēlies darbu</option>
                {data.jobs.map((j) => (
                  <option key={j} value={j}>
                    {j}
                  </option>
                ))}
              </select>

              <button className="icon-btn" onClick={addJob}>
                <Plus size={18} />
              </button>
            </div>
          </div>
        </div>

        <label>Piezīme par paveikto</label>
        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="Piemēram: sagatavota PVN deklarācija, pārbaudīti bankas konti..."
        />

        {!timer ? (
          <button className="primary-btn" onClick={startTimer}>
            <Play size={18} />
            Sākt darbu
          </button>
        ) : (
          <div className="timer-box">
            <div>
              <strong>
                {data.clients.find((c) => c.id === timer.clientId)?.name}
              </strong>
              <span>{timer.job}</span>
            </div>

            <div className="timer">
              {formatDuration(now - timer.start)}
            </div>

            <div className="timer-actions">
              <button className="secondary-btn" onClick={pauseTimer}>
                <Pause size={17} />
                Pauze
              </button>

              <button className="danger-btn" onClick={stopTimer}>
                <Square size={17} />
                Pabeigt
              </button>
            </div>
          </div>
        )}
      </section>

      <section className="card">
        <div className="section-title">
          <h2>Pēdējie ieraksti</h2>
        </div>

        {data.entries.length === 0 ? (
          <p className="muted">Pagaidām nav neviena darba laika ieraksta.</p>
        ) : (
          <div className="entries">
            {[...data.entries].reverse().slice(0, 10).map((e) => (
              <div className="entry" key={e.id}>
                <div>
                  <strong>
                    {data.clients.find((c) => c.id === e.clientId)?.name}
                  </strong>
                  <span>{e.job}</span>
                  <small>
                    {e.date} · {e.start}–{e.end}
                  </small>
                </div>

                <strong>{(e.duration / 3600000).toFixed(2)} h</strong>
              </div>
            ))}
          </div>
        )}
      </section>
    </>
  );
}

function TablePage({ data, exportExcel }) {
  const grouped = useMemo(() => {
    const result = {};

    data.clients.forEach((c) => {
      result[c.id] = {};
    });

    data.entries.forEach((e) => {
      if (!result[e.clientId]) result[e.clientId] = {};

      if (!result[e.clientId][e.date]) {
        result[e.clientId][e.date] = 0;
      }

      result[e.clientId][e.date] += e.duration;
    });

    return result;
  }, [data]);

  const dates = [
    ...new Set(data.entries.map((e) => e.date)),
  ].sort();

  function roundHours(ms) {
    return Math.round(ms / 3600000);
  }

  return (
    <section className="card">
      <div className="section-title">
        <div>
          <h2>Mēneša darba laika atskaite</h2>
          <p>Stundas pa klientiem un datumiem.</p>
        </div>

        <button className="secondary-btn" onClick={exportExcel}>
          <Download size={17} />
          Excel
        </button>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Klients</th>

              {dates.map((d) => (
                <th key={d}>{d.slice(8, 10)}.{d.slice(5, 7)}</th>
              ))}

              <th>Kopā</th>
            </tr>
          </thead>

          <tbody>
            {data.clients.map((client) => {
              const total = Object.values(
                grouped[client.id] || {}
              ).reduce((a, b) => a + b, 0);

              return (
                <tr key={client.id}>
                  <td>{client.name}</td>

                  {dates.map((date) => (
                    <td key={date}>
                      {grouped[client.id]?.[date]
                        ? `${(
                            grouped[client.id][date] / 3600000
                          ).toFixed(2)} h`
                        : "—"}
                    </td>
                  ))}

                  <td>
                    <strong>{roundHours(total)} h</strong>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="notes-box">
        <h3>Paveiktais darbs</h3>

        {data.entries.length === 0 ? (
          <p className="muted">Nav ierakstu.</p>
        ) : (
          data.entries.map((e) => (
            <div key={e.id} className="note-row">
              <strong>
                {e.date} —{" "}
                {data.clients.find((c) => c.id === e.clientId)?.name}
              </strong>
              <span>
                {e.note || e.job}
              </span>
            </div>
          ))
        )}
      </div>
    </section>
  );
}

function CalendarPage() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();

  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const firstDay = new Date(year, month, 1).getDay();

  const weekDays = ["P", "O", "T", "C", "P", "S", "Sv"];

  const holidays = {
    "01-01": "Jaungada diena",
    "05-01": "Darba svētki",
    "11-18": "Latvijas Republikas Proklamēšanas diena",
    "12-24": "Ziemassvētku vakars",
    "12-25": "Ziemassvētki",
    "12-26": "Otrie Ziemassvētki",
  };

  const cells = [];

  let start = firstDay === 0 ? 6 : firstDay - 1;

  for (let i = 0; i < start; i++) {
    cells.push(null);
  }

  for (let day = 1; day <= daysInMonth; day++) {
    cells.push(day);
  }

  return (
    <section className="card">
      <div className="section-title">
        <div>
          <h2>Kalendārs</h2>
          <p>
            {today.toLocaleDateString("lv-LV", {
              month: "long",
              year: "numeric",
            })}
          </p>
        </div>
      </div>

      <div className="calendar">
        {weekDays.map((d) => (
          <div className="calendar-head" key={d}>
            {d}
          </div>
        ))}

        {cells.map((day, index) => {
          if (!day) {
            return <div key={index} className="calendar-day empty" />;
          }

          const date = `${String(month + 1).padStart(2, "0")}-${String(
            day
          ).padStart(2, "0")}`;

          const weekday = (start + day - 1) % 7;
          const weekend = weekday >= 5;

          return (
            <div
              key={day}
              className={
                weekend || holidays[date]
                  ? "calendar-day red"
                  : "calendar-day"
              }
              title={holidays[date] || ""}
            >
              {day}
            </div>
          );
        })}
      </div>

      <div className="legend">
        <span className="legend-red" />
        Brīvdiena / svētku diena
      </div>
    </section>
  );
}

function InfoPage() {
  const topics = [
    "Nodokļu likmes",
    "PVN",
    "IIN",
    "VSAOI",
    "Minimālā alga",
    "Deklarāciju termiņi",
    "Pamatlīdzekļi",
    "Reprezentācijas izdevumi",
    "Komandējumi",
    "Gada pārskats",
    "Manas piezīmes",
  ];

  return (
    <section className="card">
      <div className="section-title">
        <div>
          <h2>Informācija</h2>
          <p>Grāmatvedības darba ātrā uzziņa.</p>
        </div>
      </div>

      <div className="info-grid">
        {topics.map((topic) => (
          <div className="info-card" key={topic}>
            <BookOpen size={19} />
            <strong>{topic}</strong>
            <span>Atvērt un rediģēt</span>
          </div>
        ))}
      </div>
    </section>
  );
}

function AnnualPage({ data, setData }) {
  const [clientId, setClientId] = useState(data.clients[0]?.id || "");

  const standardTasks = [
    "2310 atlikums pārbaudīts",
    "2310 pareizajā pusē",
    "2380 pārbaudīts",
    "Salīdzināšanas akti nosūtīti",
    "Salīdzināšanas akti saņemti",
    "Bankas atlikums pārbaudīts",
    "Kases atlikums pārbaudīts",
    "Pamatlīdzekļu inventarizācija",
    "Pamatlīdzekļu atlikumi pārbaudīti",
    "Nolietojums pārbaudīts",
    "Ieņēmumi pārbaudīti",
    "Izdevumi pārbaudīti",
    "Uzkrājumi pārbaudīti",
  ];

  const key = `${clientId}-annual`;
  const completed = data.annual[key] || {};

  function toggle(task) {
    setData((d) => ({
      ...d,
      annual: {
        ...d.annual,
        [key]: {
          ...completed,
          [task]: !completed[task],
        },
      },
    }));
  }

  return (
    <section className="card">
      <div className="section-title">
        <div>
          <h2>Gada pārskats</h2>
          <p>Pārbaudes saraksts katram klientam.</p>
        </div>
      </div>

      <select
        value={clientId}
        onChange={(e) => setClientId(e.target.value)}
      >
        {data.clients.map((c) => (
          <option value={c.id} key={c.id}>
            {c.name}
          </option>
        ))}
      </select>

      <div className="checklist">
        {standardTasks.map((task) => (
          <label className="check-row" key={task}>
            <input
              type="checkbox"
              checked={!!completed[task]}
              onChange={() => toggle(task)}
            />
            <span>{task}</span>
            {completed[task] && <Check size={18} />}
          </label>
        ))}
      </div>

      <button
        className="secondary-btn"
        onClick={() => {
          const task = prompt("Papildu uzdevums:");
          if (!task?.trim()) return;

          const current = data.annual[key] || {};

          setData((d) => ({
            ...d,
            annual: {
              ...d.annual,
              [key]: {
                ...current,
                [task.trim()]: false,
              },
            },
          }));
        }}
      >
        <Plus size={17} />
        Pievienot klienta uzdevumu
      </button>
    </section>
  );
}

function DeadlinesPage({ data, setData }) {
  const [month, setMonth] = useState(new Date().toISOString().slice(0, 7));

  const monthName = new Date(`${month}-01`).toLocaleDateString("lv-LV", {
    month: "long",
    year: "numeric",
  });

  function reportKey(clientId, report) {
    return `${month}-${clientId}-${report}`;
  }

  function toggle(clientId, report) {
    const key = reportKey(clientId, report);

    setData((d) => ({
      ...d,
      reports: {
        ...d.reports,
        [key]: !d.reports[key],
      },
    }));
  }

  return (
    <section className="card">
      <div className="section-title">
        <div>
          <h2>Termiņi un atskaites</h2>
          <p>Kontrole pa klientiem un atskaišu veidiem.</p>
        </div>
      </div>

      <input
        type="month"
        value={month}
        onChange={(e) => setMonth(e.target.value)}
      />

      <h3 style={{ marginTop: 20 }}>{monthName}</h3>

      <div className="report-list">
        {data.clients.map((client) => {
          const reports =
            client.vat === "monthly"
              ? ["PVN deklarācija"]
              : client.vat === "quarterly"
              ? ["PVN deklarācija — ceturksnis"]
              : [];

          return (
            <div className="report-client" key={client.id}>
              <strong>{client.name}</strong>

              {reports.length === 0 ? (
                <span className="muted">Nav PVN deklarācijas.</span>
              ) : (
                reports.map((report) => {
                  const done = data.reports[reportKey(client.id, report)];

                  return (
                    <label className="report-row" key={report}>
                      <input
                        type="checkbox"
                        checked={!!done}
                        onChange={() => toggle(client.id, report)}
                      />
                      <span>{report}</span>
                      {done && <Check size={18} />}
                    </label>
                  );
                })
              )}
            </div>
          );
        })}
      </div>
    </section>
  );
}

createRoot(document.getElementById("root")).render(<App />);
