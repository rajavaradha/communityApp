const emptyAttendee = () => ({ name: "", role: "" });
const emptyAction = () => ({ desc: "", resp: "", date: "" });

function DualButton({ label, shortcut, onClick, color = "emerald", type = "button" }) {
  const palette =
    color === "emerald"
      ? "bg-emerald-600 hover:bg-emerald-700 border-emerald-700"
      : "bg-teal-700 hover:bg-teal-800 border-teal-800";
  return (
    <button
      type={type}
      onClick={onClick}
      className={`min-w-[150px] rounded-md border px-3 py-1.5 text-white transition ${palette}`}
    >
      <div className="text-[10pt] font-bold leading-tight">{label}</div>
      <div className="text-[8pt] font-bold leading-tight text-sky-200">{shortcut}</div>
    </button>
  );
}

function App() {
  const [subject, setSubject] = React.useState("");
  const [date, setDate] = React.useState("");
  const [time, setTime] = React.useState("");
  const [points, setPoints] = React.useState("");
  const [attendees, setAttendees] = React.useState([emptyAttendee()]);
  const [actions, setActions] = React.useState([emptyAction()]);
  const [busy, setBusy] = React.useState(false);

  const payload = React.useMemo(
    () => ({
      subject,
      date,
      time,
      points,
      attendees,
      action_items: actions,
    }),
    [subject, date, time, points, attendees, actions]
  );

  const updateAttendee = (idx, key, value) => {
    setAttendees((prev) => prev.map((row, i) => (i === idx ? { ...row, [key]: value } : row)));
  };

  const updateAction = (idx, key, value) => {
    setActions((prev) => prev.map((row, i) => (i === idx ? { ...row, [key]: value } : row)));
  };

  const addAttendee = () => setAttendees((prev) => [...prev, emptyAttendee()]);
  const resetAttendees = () => setAttendees([emptyAttendee()]);
  const removeAttendee = (idx) => setAttendees((prev) => (prev.length === 1 ? prev : prev.filter((_, i) => i !== idx)));

  const addAction = () => setActions((prev) => [...prev, emptyAction()]);
  const resetActions = () => setActions([emptyAction()]);
  const removeAction = (idx) => setActions((prev) => (prev.length === 1 ? prev : prev.filter((_, i) => i !== idx)));

  const downloadFile = async (kind) => {
    const endpoint = kind === "word" ? "/api/export/word" : "/api/export/pdf";
    const fallbackFilename = kind === "word" ? "mom.docx" : "mom.pdf";

    setBusy(true);
    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) {
        throw new Error(`Export failed (${response.status})`);
      }
      const disposition = response.headers.get("content-disposition") || "";
      let filename = fallbackFilename;
      const filenameMatch =
        disposition.match(/filename\*=UTF-8''([^;]+)/i) ||
        disposition.match(/filename="?([^"]+)"?/i);
      if (filenameMatch && filenameMatch[1]) {
        filename = decodeURIComponent(filenameMatch[1]);
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert(err.message || "Export failed");
    } finally {
      setBusy(false);
    }
  };

  const shortcutsRef = React.useRef({});
  shortcutsRef.current = {
    addAttendee,
    resetAttendees,
    addAction,
    resetActions,
    downloadWord: () => downloadFile("word"),
    downloadPdf: () => downloadFile("pdf"),
  };

  React.useEffect(() => {
    const onGlobalShortcut = (event) => {
      if (!(event.altKey && event.shiftKey)) return;
      const h = shortcutsRef.current;
      switch (event.code) {
        case "Digit1":
          event.preventDefault();
          event.stopPropagation();
          h.addAttendee();
          break;
        case "Digit2":
          event.preventDefault();
          event.stopPropagation();
          h.resetAttendees();
          break;
        case "Digit3":
          event.preventDefault();
          event.stopPropagation();
          h.addAction();
          break;
        case "Digit4":
          event.preventDefault();
          event.stopPropagation();
          h.resetActions();
          break;
        case "KeyW":
          event.preventDefault();
          event.stopPropagation();
          h.downloadWord();
          break;
        case "KeyP":
          event.preventDefault();
          event.stopPropagation();
          h.downloadPdf();
          break;
        default:
          break;
      }
    };

    document.addEventListener("keydown", onGlobalShortcut, { capture: true });
    return () => document.removeEventListener("keydown", onGlobalShortcut, { capture: true });
  }, []);

  return (
    <main className="mx-auto my-4 max-w-6xl rounded-xl border border-emerald-200 bg-white/90 p-4 shadow-sm">
      <h1 className="mb-1 font-verdana text-[18pt] font-bold text-emerald-900">Minutes of Meeting (MOM) Creator</h1>
      <p className="mb-3 text-[10pt] text-emerald-700">React + Tailwind website UI</p>

      <section className="mb-2 rounded-lg border border-emerald-200 bg-white p-3">
        <h2 className="mb-2 text-[11pt] font-bold">Meeting Details</h2>
        <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
          <label className="flex flex-col gap-1">
            <span className="text-[10pt]">Subject</span>
            <input
              className="rounded border border-emerald-300 px-2 py-1.5"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="Project review, planning, etc."
            />
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-[10pt]">Meeting Date</span>
            <input className="rounded border border-emerald-300 px-2 py-1.5" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-[10pt]">Meeting Time</span>
            <input className="rounded border border-emerald-300 px-2 py-1.5" type="time" value={time} onChange={(e) => setTime(e.target.value)} />
          </label>
        </div>
      </section>

      <section className="mb-2 rounded-lg border border-emerald-200 bg-white p-3">
        <div className="mb-2 flex items-center justify-between gap-2">
          <h2 className="text-[11pt] font-bold">Attendees</h2>
          <div className="flex gap-[5px]">
            <DualButton label="Add" shortcut="Alt + Shift + 1" onClick={addAttendee} />
            <DualButton label="Reset" shortcut="Alt + Shift + 2" onClick={resetAttendees} />
          </div>
        </div>
        <table className="w-full border-collapse">
          <thead>
            <tr>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">S.No</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Name</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Designation / Role</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Action</th>
            </tr>
          </thead>
          <tbody>
            {attendees.map((row, idx) => (
              <tr key={`att-${idx}`}>
                <td className="border border-emerald-100 p-1.5">{idx + 1}</td>
                <td className="border border-emerald-100 p-1.5">
                  <input className="w-full rounded border border-emerald-300 px-2 py-1" value={row.name} onChange={(e) => updateAttendee(idx, "name", e.target.value)} />
                </td>
                <td className="border border-emerald-100 p-1.5">
                  <input className="w-full rounded border border-emerald-300 px-2 py-1" value={row.role} onChange={(e) => updateAttendee(idx, "role", e.target.value)} />
                </td>
                <td className="border border-emerald-100 p-1.5">
                  <button className="rounded border border-rose-300 bg-rose-50 px-2 py-1 text-rose-700" onClick={() => removeAttendee(idx)} type="button">
                    Remove
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>

      <section className="mb-2 rounded-lg border border-emerald-200 bg-white p-3">
        <h2 className="mb-2 text-[11pt] font-bold">Points Discussed</h2>
        <textarea
          className="w-full rounded border border-emerald-300 px-2 py-1.5"
          rows={8}
          value={points}
          onChange={(e) => setPoints(e.target.value)}
          placeholder="Enter points discussed"
          spellCheck
          autoCorrect="on"
          autoCapitalize="sentences"
        />
      </section>

      <section className="mb-2 rounded-lg border border-emerald-200 bg-white p-3">
        <div className="mb-2 flex items-center justify-between gap-2">
          <h2 className="text-[11pt] font-bold">Action Items</h2>
          <div className="flex gap-[5px]">
            <DualButton label="Add" shortcut="Alt + Shift + 3" onClick={addAction} />
            <DualButton label="Reset" shortcut="Alt + Shift + 4" onClick={resetActions} />
          </div>
        </div>
        <table className="w-full border-collapse">
          <thead>
            <tr>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Action Item</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Responsibility</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Timeline</th>
              <th className="border border-emerald-200 bg-emerald-50 p-1.5 text-left text-[10pt]">Action</th>
            </tr>
          </thead>
          <tbody>
            {actions.map((row, idx) => (
              <tr key={`action-${idx}`}>
                <td className="border border-emerald-100 p-1.5">
                  <input className="w-full rounded border border-emerald-300 px-2 py-1" value={row.desc} onChange={(e) => updateAction(idx, "desc", e.target.value)} />
                </td>
                <td className="border border-emerald-100 p-1.5">
                  <input className="w-full rounded border border-emerald-300 px-2 py-1" value={row.resp} onChange={(e) => updateAction(idx, "resp", e.target.value)} />
                </td>
                <td className="border border-emerald-100 p-1.5">
                  <input className="w-full rounded border border-emerald-300 px-2 py-1" value={row.date} onChange={(e) => updateAction(idx, "date", e.target.value)} />
                </td>
                <td className="border border-emerald-100 p-1.5">
                  <button className="rounded border border-rose-300 bg-rose-50 px-2 py-1 text-rose-700" onClick={() => removeAction(idx)} type="button">
                    Remove
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>

      <section className="rounded-lg border border-emerald-200 bg-white p-3">
        <h2 className="mb-2 text-[11pt] font-bold">Download Documents</h2>
        <div className="flex gap-[5px]">
          <DualButton label="Word" shortcut="Alt + Shift + W" onClick={() => downloadFile("word")} color="teal" />
          <DualButton label="PDF" shortcut="Alt + Shift + P" onClick={() => downloadFile("pdf")} color="teal" />
        </div>
        {busy && <p className="mt-2 text-[10pt] text-emerald-700">Preparing file...</p>}
      </section>
    </main>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
