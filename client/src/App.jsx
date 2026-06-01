import { useEffect, useMemo, useState } from "react";
import "./App.css";

const AUTH_STORAGE_KEY = "workflow-app-auth";

function buildInitialValues(workflow) {
  const values = {};
  for (const field of workflow?.fields || []) {
    values[field.name] = field.type === "checkbox" ? Boolean(field.defaultValue) : field.defaultValue ?? "";
  }
  return values;
}

function loadStoredAuth() {
  if (typeof window === "undefined") return null;

  try {
    const raw = window.localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed?.token || !parsed?.email) return null;
    return parsed;
  } catch {
    return null;
  }
}

function persistAuth(auth) {
  if (typeof window === "undefined") return;

  if (!auth) {
    window.localStorage.removeItem(AUTH_STORAGE_KEY);
    return;
  }

  window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(auth));
}

async function readResponse(response) {
  const raw = await response.text();
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return { ok: false, error: raw };
  }
}

function isFileFieldPresent(workflow) {
  return (workflow?.fields || []).some((field) => field.type === "file");
}

function renderEntry(entry) {
  return typeof entry === "string" ? entry : JSON.stringify(entry, null, 2);
}

function fileExample(field) {
  const label = String(field.label || "").toLowerCase();

  if (label.includes("barcode")) {
    return 'Example: TXT/CSV/XLSX with one barcode per line or first-column values like "8690604534746".';
  }

  if (label.includes("sku")) {
    return 'Example: TXT/CSV/XLSX with one SKU per line or first-column values like "41000023-00".';
  }

  return "Example: TXT/CSV/XLSX with one value per line or in the first column of the first sheet.";
}

export default function App() {
  const [auth, setAuth] = useState(() => loadStoredAuth());
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [workflows, setWorkflows] = useState([]);
  const [selectedId, setSelectedId] = useState("");
  const [formValues, setFormValues] = useState({});
  const [downloads, setDownloads] = useState([]);
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");

  function applyLogout() {
    setAuth(null);
    persistAuth(null);
    setLoginPassword("");
    setWorkflows([]);
    setSelectedId("");
    setFormValues({});
    setDownloads([]);
    setLogs([]);
    setError("");
  }

  async function authFetch(input, init = {}) {
    const response = await fetch(input, {
      ...init,
      headers: {
        ...(init.headers || {}),
        "x-auth-token": auth?.token || "",
      },
    });

    if (response.status === 401) {
      applyLogout();
    }

    return response;
  }

  useEffect(() => {
    if (!auth?.token) {
      setLoading(false);
      return undefined;
    }

    let active = true;

    async function load() {
      try {
        setLoading(true);
        setError("");
        const response = await authFetch("/api/workflows");
        const payload = await readResponse(response);
        if (!response.ok || !payload?.ok) {
          throw new Error(payload?.error || `Failed to load workflows (${response.status}).`);
        }
        if (!active) return;
        setWorkflows(payload.workflows || []);
        if (payload.workflows?.length) {
          setSelectedId(payload.workflows[0].id);
          setFormValues(buildInitialValues(payload.workflows[0]));
        }
      } catch (err) {
        if (active) setError(err.message || "Failed to load workflows.");
      } finally {
        if (active) setLoading(false);
      }
    }

    load();
    return () => {
      active = false;
    };
  }, [auth?.token]);

  const selectedWorkflow = useMemo(
    () => workflows.find((workflow) => workflow.id === selectedId) ?? null,
    [workflows, selectedId],
  );

  function appendLog(value) {
    setLogs((current) => [
      {
        id: `${Date.now()}-${current.length}`,
        at: new Date().toLocaleTimeString(),
        value,
      },
      ...current,
    ]);
  }

  function handleSelect(workflow) {
    setSelectedId(workflow.id);
    setFormValues(buildInitialValues(workflow));
    setDownloads([]);
  }

  function updateField(name, value) {
    setFormValues((current) => ({ ...current, [name]: value }));
  }

  function buildRequest(workflow) {
    if (isFileFieldPresent(workflow)) {
      const formData = new FormData();
      for (const field of workflow.fields || []) {
        const value = formValues[field.name];
        if (field.type === "file") {
          if (value instanceof File) formData.append(field.name, value);
          continue;
        }
        formData.append(field.name, field.type === "checkbox" ? (value ? "1" : "0") : value ?? "");
      }
      return { body: formData };
    }

    const payload = {};
    for (const field of workflow.fields || []) {
      payload[field.name] = formValues[field.name];
    }
    return {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    };
  }

  async function runWorkflow() {
    if (!selectedWorkflow) return;

    setRunning(true);
    setDownloads([]);
    appendLog({
      action: "start",
      workflow: selectedWorkflow.id,
      inputs: Object.fromEntries(
        Object.entries(formValues).map(([key, value]) => [key, value instanceof File ? value.name : value]),
      ),
    });

    try {
      const response = await authFetch(selectedWorkflow.endpoint, {
        method: selectedWorkflow.method || "POST",
        ...buildRequest(selectedWorkflow),
      });
      const payload = await readResponse(response);
      if (!response.ok || payload?.ok === false) {
        throw new Error(payload?.error || `Request failed (${response.status}).`);
      }
      if (!payload) {
        throw new Error("Server returned an empty response.");
      }

      setDownloads(payload.downloads || []);
      appendLog(payload.summary || payload);
    } catch (err) {
      appendLog(`Request failed: ${err.message || String(err)}`);
    } finally {
      setRunning(false);
    }
  }

  function downloadFile(file) {
    const blob = new Blob([file.content], { type: file.mimeType || "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = file.filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  async function handleLogin(event) {
    event.preventDefault();

    try {
      setLoginLoading(true);
      setLoginError("");

      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          email: loginEmail,
          password: loginPassword,
        }),
      });
      const payload = await readResponse(response);
      if (!response.ok || !payload?.ok || !payload?.token) {
        throw new Error(payload?.error || "Login failed.");
      }

      const nextAuth = {
        token: payload.token,
        email: payload.user?.email || loginEmail,
      };
      setAuth(nextAuth);
      persistAuth(nextAuth);
      setLoginPassword("");
    } catch (err) {
      setLoginError(err.message || "Login failed.");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    try {
      if (auth?.token) {
        await fetch("/api/auth/logout", {
          method: "POST",
          headers: {
            "x-auth-token": auth.token,
          },
        });
      }
    } catch {
      // Logout should still clear local state even if the request fails.
    } finally {
      applyLogout();
    }
  }

  if (!auth?.token) {
    return (
      <div className="login-shell">
        <section className="login-card">
          <p className="eyebrow">Secure Access</p>
          <h1>Workflow Login</h1>
          <p className="login-copy">Use one of the approved Flormar emails and the shared password to access the workflow app.</p>

          <form className="login-form" onSubmit={handleLogin}>
            <label className="field">
              <span>Email</span>
              <input
                type="email"
                value={loginEmail}
                placeholder="name@flormar.ma"
                onChange={(event) => setLoginEmail(event.target.value)}
                required
              />
            </label>

            <label className="field">
              <span>Password</span>
              <input
                type="password"
                value={loginPassword}
                placeholder="Password"
                onChange={(event) => setLoginPassword(event.target.value)}
                required
              />
            </label>

            {loginError ? <p className="error-text">{loginError}</p> : null}

            <button type="submit" className="primary-button" disabled={loginLoading}>
              {loginLoading ? "Signing In..." : "Sign In"}
            </button>
          </form>

          <div className="login-help">
            <p className="section-title">Allowed Emails</p>
            <p>a.elhabchi@flormar.ma</p>
            <p>y.bajou@flormar.ma</p>
            <p>a.chafa@flormar.ma</p>
          </div>
        </section>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <p className="eyebrow">WooCommerce Ops</p>
          <h1>Workflow App</h1>
          <p className="subtle">A clean browser frontend for your WooCommerce operations.</p>
          <p className="subtle signed-in-as">Signed in as {auth.email}</p>
        </div>

        <div className="sidebar-section">
          <p className="section-title">Workflows</p>
          {loading ? <p className="subtle">Loading workflows...</p> : null}
          {error ? <p className="error-text">{error}</p> : null}
          <div className="workflow-list">
            {workflows.map((workflow) => (
              <button
                key={workflow.id}
                type="button"
                className={`workflow-card ${selectedId === workflow.id ? "selected" : ""}`}
                onClick={() => handleSelect(workflow)}
              >
                <strong>{workflow.title}</strong>
                <span>{workflow.category}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="sidebar-section">
          <button type="button" className="secondary-button" onClick={handleLogout}>
            Logout
          </button>
        </div>
      </aside>

      <main className="main-panel">
        {selectedWorkflow ? (
          <>
            <section className="hero-panel">
              <div>
                <p className="eyebrow">{selectedWorkflow.category}</p>
                <h2>{selectedWorkflow.title}</h2>
                <p className="hero-copy">{selectedWorkflow.description}</p>
              </div>
              <div className="hero-meta">
                <span className="meta-label">Endpoint</span>
                <p>{selectedWorkflow.endpoint}</p>
              </div>
            </section>

            <section className="content-grid">
              <div className="panel">
                <div className="panel-head">
                  <h3>Inputs</h3>
                  <p>Fill the form and run the workflow from the browser.</p>
                </div>

                <div className="form-grid">
                  {(selectedWorkflow.fields || []).map((field) => (
                    <label key={field.name} className="field">
                      <span>{field.label}</span>

                      {field.type === "textarea" ? (
                        <textarea
                          value={formValues[field.name] ?? ""}
                          placeholder={field.placeholder || ""}
                          onChange={(event) => updateField(field.name, event.target.value)}
                        />
                      ) : null}

                      {["text", "date", "number"].includes(field.type) ? (
                        <input
                          type={field.type}
                          value={formValues[field.name] ?? ""}
                          placeholder={field.placeholder || ""}
                          onChange={(event) => updateField(field.name, event.target.value)}
                        />
                      ) : null}

                      {field.type === "select" ? (
                        <select
                          value={formValues[field.name] ?? ""}
                          onChange={(event) => updateField(field.name, event.target.value)}
                        >
                          {(field.options || []).map((option) => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                      ) : null}

                      {field.type === "file" ? (
                        <>
                          <input
                            type="file"
                            accept={field.accept}
                            onChange={(event) => updateField(field.name, event.target.files?.[0] || null)}
                          />
                          <small>{fileExample(field)}</small>
                        </>
                      ) : null}

                      {field.type === "checkbox" ? (
                        <div className="checkbox-row">
                          <input
                            type="checkbox"
                            checked={Boolean(formValues[field.name])}
                            onChange={(event) => updateField(field.name, event.target.checked)}
                          />
                          <span>{field.help || "Enabled"}</span>
                        </div>
                      ) : null}

                      {field.help && field.type !== "checkbox" ? <small>{field.help}</small> : null}
                    </label>
                  ))}
                </div>

                <div className="actions">
                  <button type="button" className="primary-button" disabled={running} onClick={runWorkflow}>
                    {running ? "Running..." : "Run Workflow"}
                  </button>
                </div>
              </div>

              <div className="panel">
                <div className="panel-head">
                  <h3>Downloads</h3>
                  <p>Generated files appear here.</p>
                </div>

                {downloads.length ? (
                  <div className="download-list">
                    {downloads.map((file) => (
                      <button
                        key={file.filename}
                        type="button"
                        className="download-card"
                        onClick={() => downloadFile(file)}
                      >
                        <strong>{file.filename}</strong>
                        <span>Download</span>
                      </button>
                    ))}
                  </div>
                ) : (
                  <div className="empty-state compact">
                    <p>No downloads yet.</p>
                    <span>Exports will appear here.</span>
                  </div>
                )}

                <div className="panel-head log-head">
                  <h3>Execution Log</h3>
                  <p>Recent summaries and errors.</p>
                </div>

                <div className="log-list">
                  {logs.length ? (
                    logs.map((entry) => (
                      <article key={entry.id} className="log-entry">
                        <span>{entry.at}</span>
                        <pre>{renderEntry(entry.value)}</pre>
                      </article>
                    ))
                  ) : (
                    <div className="empty-state">
                      <p>No runs yet.</p>
                      <span>Run any workflow to see the result here.</span>
                    </div>
                  )}
                </div>
              </div>
            </section>
          </>
        ) : null}
      </main>
    </div>
  );
}
