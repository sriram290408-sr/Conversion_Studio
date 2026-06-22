(() => {
    const { useState: c, useRef: le, useEffect: Ye } = React,
        N = (path => {
            let p = path;
            if (!p.startsWith("/api/")) {
                p = p.startsWith("/") ? `/api${p}` : `/api/${p}`;
            }
            return window.apiURL(p);
        }),
        LN = (path => {
            let p = path;
            if (!p.startsWith("/api/")) {
                p = p.startsWith("/") ? `/api${p}` : `/api/${p}`;
            }
            return window.liveApiURL(p);
        }),
        u = ({ id: o, size: v = 16, cls: r = "" }) => React.createElement("svg", { width: v, height: v, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: "1.8", strokeLinecap: "round", strokeLinejoin: "round", className: r, style: { flexShrink: 0 } }, { bolt: React.createElement(React.Fragment, null, React.createElement("path", { d: "M13 2L4.5 13h6.5l-1 7 8.5-11H12l1-7z" })), upload: React.createElement(React.Fragment, null, React.createElement("path", { d: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" }), React.createElement("polyline", { points: "17 8 12 3 7 8" }), React.createElement("line", { x1: "12", y1: "3", x2: "12", y2: "15" })), img: React.createElement(React.Fragment, null, React.createElement("rect", { x: "3", y: "3", width: "18", height: "18", rx: "2" }), React.createElement("circle", { cx: "8.5", cy: "8.5", r: "1.5" }), React.createElement("polyline", { points: "21 15 16 10 5 21" })), check: React.createElement(React.Fragment, null, React.createElement("polyline", { points: "20 6 9 17 4 12" })), dl: React.createElement(React.Fragment, null, React.createElement("path", { d: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" }), React.createElement("polyline", { points: "7 10 12 15 17 10" }), React.createElement("line", { x1: "12", y1: "15", x2: "12", y2: "3" })), x: React.createElement(React.Fragment, null, React.createElement("line", { x1: "18", y1: "6", x2: "6", y2: "18" }), React.createElement("line", { x1: "6", y1: "6", x2: "18", y2: "18" })), eye: React.createElement(React.Fragment, null, React.createElement("path", { d: "M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" }), React.createElement("circle", { cx: "12", cy: "12", r: "3" })), arrow: React.createElement(React.Fragment, null, React.createElement("line", { x1: "5", y1: "12", x2: "19", y2: "12" }), React.createElement("polyline", { points: "12 5 19 12 12 19" })) }[o]), _ = ({ n: o, label: v, st: r }) => React.createElement("div", { className: `step ${r === "running" ? "active" : ""} ${r === "done" ? "done" : ""}` }, React.createElement("div", { className: "step-node" }, r === "done" ? React.createElement(u, { id: "check", size: 10 }) : React.createElement("span", { className: "step-num" }, o)), React.createElement("span", { className: "step-label" }, v)), as = ({ data: o, tab: v }) => { var p, C, z, B, U, T, q; const r = o == null ? void 0 : o.chunks_preview; if (!o || !r) return React.createElement("div", { className: "empty" }, React.createElement("div", { className: "empty-icon" }, React.createElement(u, { id: "eye", size: 28 })), React.createElement("p", { className: "empty-label" }, "Awaiting conversion"), React.createElement("p", { className: "empty-desc" }, "Upload the PBIX, dashboard screenshot, and TMDL metadata to analyze the report, preview the report structure, and prepare the interactive live Excel connection."), React.createElement("div", { className: "empty-arrow" }, React.createElement(u, { id: "arrow", size: 12 }), React.createElement("span", null, "Select all three required files from the left panel to begin"))); if (v === "raw") return React.createElement("pre", { className: "raw-pre preview-surface" }, ((p = r.raw_json_preview) == null ? void 0 : p.text) || JSON.stringify(o, null, 2)); const E = { analysis: (C = r.metadata_analysis_preview) == null ? void 0 : C.html, tables: (z = r.tables_preview) == null ? void 0 : z.html, relationships: (B = r.relationships_preview) == null ? void 0 : B.html, formulas: (U = r.formulas_preview) == null ? void 0 : U.html, visuals: (T = r.visuals_preview) == null ? void 0 : T.html, descriptions: (q = r.visual_descriptions_preview) == null ? void 0 : q.html }; return React.createElement("div", { className: "preview-surface", dangerouslySetInnerHTML: { __html: E[v] || `<div class="empty-tab">No ${v} data found</div>` } }) }, ts = () => {
            var xe, Pe, Ce, Le, Se, De, Ee, Be, Te; const [o, v] = c(null), [r, E] = c(null), [p, C] = c(null), [z, B] = c(null), [U, T] = c(!1), [q, ee] = c(!1), [$e, se] = c(!1), [L, oe] = c(!1), [w, re] = c(!1), [a, A] = c(null), [ce, W] = c("analysis"), [de, m] = c(null), [V, I] = c({ a: "idle", b: "idle", c: "idle", d: "idle" }), [l, y] = c(null), [M, S] = c(!1), [h, ve] = c(null), [ns, pe] = c(null), [X, H] = c(null), [Re, me] = c(0), [is, ae] = c([]), [ue, F] = c(null), [J, G] = c(null), [k, Oe] = c(() => localStorage.getItem("conversion-studio-mode") || "dark"),
                [capabilities, setCapabilities] = c({ standardConversion: true, liveConnect: false }),
                [backendStatus, setBackendStatus] = c("Checking status...");

            Ye(() => {
                let cancelled = false;

                const checkBackends = async () => {
                    let cloudReady = false;
                    let agentReady = false;
                    let agentHealth = null;

                    try {
                        const cloudResponse = await fetch(window.apiURL("/api/health"), { cache: "no-store" });
                        cloudReady = cloudResponse.ok;
                    } catch (_) {
                        cloudReady = false;
                    }

                    try {
                        const agentResponse = await fetch(window.liveApiURL("/api/health"), {
                            cache: "no-store",
                            headers: { "X-Conversion-Studio-Client": "vercel-frontend" }
                        });
                        if (agentResponse.ok) {
                            agentHealth = await agentResponse.json();
                            agentReady =
                                agentHealth.platform === "Windows" &&
                                agentHealth.live_connect_available === true;
                        }
                    } catch (_) {
                        agentReady = false;
                    }

                    if (cancelled) return;

                    setCapabilities({
                        standardConversion: cloudReady || window.IS_LOCAL_WINDOWS_MODE,
                        liveConnect: agentReady
                    });

                    if (agentReady) {
                        setBackendStatus(
                            window.IS_LOCAL_WINDOWS_MODE
                                ? "Local Windows mode — Live Connect ready"
                                : "Cloud online · Windows agent connected"
                        );
                    } else if (cloudReady) {
                        setBackendStatus("Cloud online · Windows agent offline");
                    } else {
                        setBackendStatus("Backend unavailable");
                    }
                };

                checkBackends();
                const timer = window.setInterval(checkBackends, 10000);
                return () => {
                    cancelled = true;
                    window.clearInterval(timer);
                };
            }, []);

            Ye(() => { localStorage.setItem("conversion-studio-mode", k), document.documentElement.style.colorScheme = k }, [k]); const te = e => e && (e.state || e.status) || null, ne = new Set(["completed_live", "live_conversion_failed", "cancelled", "error"]), fe = new Set(["waiting_for_user_connection", "connection_not_detected", "semantic_model_mismatch"]), be = (e, s) => { const n = setInterval(async () => { var i; try { const t = await fetch(LN(`/live-connect/${e}/status`)); if (!t.ok) return; const d = await t.json(), f = te(d); f && y(f), d.pivot_tables_created !== void 0 && me(d.pivot_tables_created), d.semantic_match_score !== void 0 && H(d.semantic_match_score), (i = d.warnings) != null && i.length && ae(d.warnings), d.error_message ? F(`${d.error_message} [Stage: ${d.error_stage || 'unknown'}, Trace: ${d.error_trace_id || 'N/A'}]`) : (d.message && F(d.message)), (ne.has(f) || fe.has(f)) && (clearInterval(n), G(null), S(!1), s && s(f, d)) } catch (t) { console.error("Status poll error:", t) } }, 1500); return G(n), n }, ze = async () => {
                if (!capabilities.liveConnect) {
                    F(
                        "Windows agent not detected. Start start-local-agent.bat on this laptop, " +
                        "keep the terminal open, and try again."
                    );
                    y(null);
                    S(false);
                    return;
                }

                if (!o || !r || !p) {
                    F("PBIX, dashboard screenshot, and TMDL metadata are required for Live Connect.");
                    return;
                }

                if (M || l && !ne.has(l)) return;

                S(true);
                F(null);
                y("preparing_local_session");

                try {
                    // The cloud upload session cannot be reused by the Windows process.
                    // Upload the same selected source files directly to the local agent.
                    const localForm = new FormData();
                    localForm.append("file", o);
                    localForm.append("screenshot", r);
                    localForm.append("tmdl_metadata", p);

                    const localUploadResponse = await fetch(LN("/upload"), {
                        method: "POST",
                        body: localForm
                    });

                    if (!localUploadResponse.ok) {
                        const text = await localUploadResponse.text();
                        throw new Error(text || `Local preparation failed (${localUploadResponse.status})`);
                    }

                    const localUpload = await localUploadResponse.json();
                    if (!localUpload.session_id) {
                        throw new Error("The Windows agent did not return a local upload session ID.");
                    }

                    y("excel_launching");
                    const response = await fetch(LN("/live-connect/start"), {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ session_id: localUpload.session_id })
                    });

                    if (!response.ok) {
                        const text = await response.text();
                        throw new Error(text || `Failed to start local session (${response.status})`);
                    }

                    const payload = await response.json();
                    const state = te(payload);
                    ve(payload.session_id);
                    y(state || "excel_launching");
                    pe(payload.message);
                    if (!fe.has(state) && !ne.has(state)) {
                        be(payload.session_id, null);
                    }
                } catch (error) {
                    console.error(error);
                    F(error.message || "Failed to start the Windows Live Connect workflow.");
                    y("live_conversion_failed");
                } finally {
                    S(false);
                }
            }, Ae = async () => { if (h && !M) { S(!0), F(null); try { const e = await fetch(LN(`/live-connect/${h}/continue`), { method: "POST" }); if (!e.ok) { const i = await e.text(); throw new Error(i || `Failed to continue session (${e.status})`) } const s = await e.json(), n = te(s); n && y(n), be(h, (i, t) => { var d; i === "completed_live" ? (t.semantic_match_score !== void 0 && H(t.semantic_match_score), (d = t.warnings) != null && d.length && ae(t.warnings), je(h)) : i === "semantic_model_mismatch" && t.semantic_match_score !== void 0 && H(t.semantic_match_score) }) } catch (e) { console.error(e), F(e.message || "Failed to build live connection."), y("live_conversion_failed"), S(!1) } } }, Me = async () => { if (h) { S(!0); try { if (J && (clearInterval(J), G(null)), !(await fetch(LN(`/live-connect/${h}/cancel`), { method: "POST" })).ok) throw new Error; y("cancelled") } catch (e) { console.error(e), F("Failed to cancel Live Connect session."), y("live_conversion_failed") } finally { S(!1) } } }, K = () => { J && (clearInterval(J), G(null)), y(null), ve(null), pe(null), H(null), me(0), ae([]), F(null) }, je = async e => { try { const s = await fetch(LN(`/live-connect/${e}/report`)); if (s.ok) { const n = await s.json(); A(i => ({ ...i, chunks_preview: n.chunks_preview || i.chunks_preview, summary: n.summary || i.summary })) } } catch (s) { console.error("Failed to fetch live session report:", s) } }, Ue = () => { h && window.open(LN(`/live-connect/${h}/download`), "_blank") }, qe = () => { h && window.open(LN(`/live-connect/${h}/report`), "_blank") }, $ = le(null), R = le(null), x = le(null), b = (a == null ? void 0 : a.summary) || {}, O = (a == null ? void 0 : a.chunks_preview) || {}, g = { analysis: ((xe = O.metadata_analysis_preview) == null ? void 0 : xe.count) || ((Ce = (Pe = a == null ? void 0 : a.metadata_analysis) == null ? void 0 : Pe.overall_counts) == null ? void 0 : Ce.total_visuals) || 0, tables: ((Le = O.tables_preview) == null ? void 0 : Le.count) || b.total_model_tables || b.total_tables || b.tables || 0, relationships: ((Se = O.relationships_preview) == null ? void 0 : Se.count) || b.total_relationships || b.relationships || 0, formulas: ((De = O.formulas_preview) == null ? void 0 : De.count) || b.total_formulas || b.measures || 0, visuals: ((Ee = O.visuals_preview) == null ? void 0 : Ee.count) || b.total_visuals || b.visuals || 0, descriptions: ((Te = (Be = O.visual_descriptions_preview) == null ? void 0 : Be.records) == null ? void 0 : Te.length) || b.total_visuals || b.visuals || 0 }, Y = () => I({ a: "idle", b: "idle", c: "idle", d: "idle" }), D = () => { A(null), W("analysis"), Y(), K() }, he = e => { if (e) { if (!e.name.toLowerCase().endsWith(".pbix")) { m("Only .pbix files are accepted."), v(null), D(), $.current && ($.current.value = ""); return } v(e), m(null), D() } }, ge = e => { if (!e) return; const s = e.name.toLowerCase().split(".").pop(); if (!["png", "jpg", "jpeg", "webp"].includes(s)) { m("Screenshot must be .png, .jpg, .jpeg, or .webp."); return } E(e), m(null), D(); const n = new FileReader; n.onloadend = () => B(n.result), n.readAsDataURL(e) }, we = e => { if (!e) return; const s = e.name.toLowerCase(); if (![".tmdl", ".txt", ".json"].some(i => s.endsWith(i))) { m("Model metadata must be .tmdl, .txt, or .json."), C(null), D(), x.current && (x.current.value = ""); return } C(e), m(null), D() }, We = e => { e == null || e.stopPropagation(), E(null), B(null), D(), R.current && (R.current.value = "") }, Ve = e => { e == null || e.stopPropagation(), C(null), D(), x.current && (x.current.value = "") }, Xe = async () => { if (!o || !r || !p) { m("PBIX, dashboard screenshot, and TMDL metadata are required."); return } oe(!0), m(null), A(null), Y(); const e = s => new Promise(n => setTimeout(n, s)); try { I(t => ({ ...t, a: "running" })), await e(200); const s = new FormData; s.append("file", o), r && s.append("screenshot", r), p && s.append("tmdl_metadata", p); const n = await fetch(N("/upload"), { method: "POST", body: s }); if (!n.ok) { const t = await n.text(); let d = `Upload failed (${n.status})`; try { const f = JSON.parse(t); f.detail && (d = typeof f.detail == "string" ? f.detail : JSON.stringify(f.detail)) } catch (f) { t && (d = t) } throw new Error(d) } I(t => ({ ...t, a: "done", b: "running" })), await e(200); const i = await n.json(); if (!i || !i.download_url && !i.session_id) throw new Error("Backend did not return a download URL."); I(t => ({ ...t, b: "done", c: "running" })), await e(200), I(t => ({ ...t, c: "done", d: "running" })), await e(200), I(t => ({ ...t, d: "done" })), A(i), W("analysis") } catch (s) { m(s.message || "Processing failed."), Y() } finally { oe(!1) } }, He = () => { v(null), E(null), C(null), B(null), A(null), m(null), W("analysis"), Y(), K(), $.current && ($.current.value = ""), R.current && (R.current.value = ""), x.current && (x.current.value = "") }, ye = async () => { if (!(a != null && a.session_id)) { m("Run the conversion before downloading the preview PDF."); return } re(!0), m(null); try { if (a.pdf_status && a.pdf_status.download_ready === !1) throw new Error(a.pdf_status.error || "The server-side preview PDF is not available."); const e = a.preview_download_url || `/download-preview/${a.session_id}`, s = e.includes("?") ? "&" : "?", n = `${N(e)}${s}_=${Date.now()}`, i = await fetch(n, { method: "GET", cache: "no-store", headers: { Accept: "application/pdf" } }); if (!i.ok) { const ie = await i.text(); let Q = `Preview PDF download failed (${i.status}).`; try { const Z = JSON.parse(ie), P = Z == null ? void 0 : Z.detail; Q = (P == null ? void 0 : P.reason) || (P == null ? void 0 : P.message) || (typeof P == "string" ? P : Q) } catch (Z) { ie && (Q = ie) } throw new Error(Q) } const t = (i.headers.get("content-type") || "").toLowerCase(); if (!t.includes("application/pdf")) throw new Error(`Backend returned ${t || "an unknown content type"} instead of PDF.`); const d = new Uint8Array(await i.arrayBuffer()); if (d.length < 1e3) throw new Error(`Downloaded PDF is empty or incomplete (${d.length} bytes).`); if (String.fromCharCode(...d.slice(0, 5)) !== "%PDF-") throw new Error("Downloaded file is not a valid PDF."); const Ke = new Blob([d], { type: "application/pdf" }), Ie = URL.createObjectURL(Ke), j = document.createElement("a"); j.href = Ie, j.download = a.preview_filename || "powerbi_chunk_visualizer_preview.pdf", document.body.appendChild(j), j.click(), j.remove(), window.setTimeout(() => URL.revokeObjectURL(Ie), 3e3) } catch (e) { console.error("Preview PDF download failed:", e), m((e == null ? void 0 : e.message) || "Preview PDF download failed.") } finally { re(!1) } }, Ne = L || w, _e = !!(o || r || p), Je = L ? "Processing…" : w ? "Preparing PDF…" : a ? "Completed" : _e ? "Files selected" : "Idle", ke = L || w ? "processing" : a ? "success" : _e ? "uploaded" : "idle", Ge = [{ id: "analysis", label: "Analysis", n: g.analysis }, { id: "tables", label: "Tables", n: g.tables }, { id: "relationships", label: "Relations", n: g.relationships }, { id: "formulas", label: "Formulas", n: g.formulas }, { id: "visuals", label: "Visuals", n: g.visuals }, { id: "descriptions", label: "Descriptions", n: g.descriptions }, { id: "raw", label: "Raw JSON", n: null }];

            const renderLiveConnectCard = () => {
                if (!a || a.has_live_excel_workbook) return null;
                if (!capabilities.liveConnect) {
                    return React.createElement("div", { className: "card span2" },
                        React.createElement("div", { className: "card-head" },
                            React.createElement("span", { className: "card-title" }, "Live Power BI Connection"),
                            React.createElement("span", { className: "badge-opt" }, "Windows Agent Offline")
                        ),
                        React.createElement("div", { className: "card-body" },
                            React.createElement("div", { className: "ready-note", style: { background: "rgba(201, 168, 76, 0.1)", color: "var(--gold2)", marginBottom: 12, borderLeft: "3px solid var(--gold)" } },
                                React.createElement("strong", null, "Start the local Windows agent to enable Excel COM."),
                                React.createElement("div", { style: { marginTop: 5 } }, "Run start-local-agent.bat from the project folder and keep the terminal open. This Vercel page will detect it automatically.")
                            ),
                            React.createElement("button", {
                                className: "btn-clear",
                                type: "button",
                                onClick: () => window.open(window.LOCAL_AGENT_API_URL + "/api/system-check", "_blank")
                            }, "Check Windows Agent")
                        )
                    );
                }
                return React.createElement("div", { className: "card span2" },
                    React.createElement("div", { className: "card-head" },
                        React.createElement("span", { className: "card-title" }, "Live Power BI Connection"),
                        React.createElement("span", { className: `badge-${l === "completed_live" ? "req" : "opt"}` }, l === "completed_live" ? "LIVE Connected" : "Interactive")
                    ),
                    React.createElement("div", { className: "card-body" },
                        !l && React.createElement(React.Fragment, null,
                            React.createElement("p", { className: "helper-note", style: { marginBottom: 12 } }, "Metadata analysis is complete. Open Excel, select the published Power BI semantic model, and let the application build the live-connected dashboard."),
                            React.createElement("button", { className: "btn-run", onClick: ze, disabled: Ne || M },
                                React.createElement(u, { id: "bolt", size: 13 }), " Open Excel & Connect Live"
                            )
                        ),
                        l === "preparing_local_session" && React.createElement("div", { style: { textAlign: "center", padding: "12px 0" } },
                            React.createElement("span", { className: "spin", style: { display: "inline-block", marginBottom: 8 } }),
                            React.createElement("div", { className: "drop-label" }, "Preparing files on the Windows agent...")
                        ),
                        l === "excel_launching" && React.createElement("div", { style: { textAlign: "center", padding: "12px 0" } },
                            React.createElement("span", { className: "spin", style: { display: "inline-block", marginBottom: 8 } }),
                            React.createElement("div", { className: "drop-label" }, "Launching Excel visibly...")
                        ),
                        (l === "waiting_for_user_connection" || l === "connection_not_detected" || l === "semantic_model_mismatch") && React.createElement("div", null,
                            React.createElement("div", { className: "ready-note", style: { background: "rgba(201, 168, 76, 0.1)", color: "var(--gold2)", marginBottom: 12, borderLeft: "3px solid var(--gold)" } },
                                React.createElement("strong", null, "Excel is open. Please:"),
                                React.createElement("ol", { style: { paddingLeft: 16, marginTop: 4, fontSize: 12, lineHeight: 1.5 } },
                                    React.createElement("li", null, "Sign in to your Microsoft account in Excel if prompted."),
                                    React.createElement("li", null, "Go to ", React.createElement("strong", null, "Insert > PivotTable > From Power BI"), "."),
                                    React.createElement("li", null, "Select your published Power BI semantic model."),
                                    React.createElement("li", null, "Once the empty PivotTable is created, click ", React.createElement("strong", null, "Continue"), " below.")
                                )
                            ),
                            l === "connection_not_detected" && React.createElement("div", { className: "error-box", style: { marginBottom: 12 } }, "No Power BI PivotTable connection detected in Excel. Please try inserting the PivotTable and clicking Continue again."),
                            l === "semantic_model_mismatch" && React.createElement("div", { className: "error-box", style: { marginBottom: 12 } }, "The selected Power BI model does not match this PBIX metadata (Match Score: ", X !== null ? (X * 100).toFixed(0) : "0", "%). Please delete the PivotTable, insert a PivotTable from the correct model, and click Continue again."),
                            React.createElement("div", { style: { display: "flex", gap: 8 } },
                                React.createElement("button", { className: "btn-run", onClick: Ae, style: { flex: 1 }, disabled: M },
                                    M ? React.createElement("span", { className: "spin" }) : React.createElement(u, { id: "bolt", size: 13 }), "Continue"
                                ),
                                React.createElement("button", { className: "btn-clear", onClick: Me, style: { border: "1px solid var(--rose)", color: "var(--rose)" } }, "Cancel")
                            )
                        ),
                        ["detecting_connection", "connection_detected", "validating_semantic_model", "building", "refreshing", "saving", "verifying"].includes(l) && React.createElement("div", { style: { padding: "8px 0" } },
                            React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 12 } },
                                React.createElement("span", { className: "spin" }),
                                React.createElement("div", { className: "drop-label", style: { textTransform: "capitalize" } }, l.replace(/_/g, " "), "...")
                            ),
                            React.createElement("div", { className: "steps" },
                                React.createElement(_, { n: "01", label: "Detect Power BI connection", st: l === "detecting_connection" ? "running" : "done" }),
                                React.createElement(_, { n: "02", label: "Validate semantic model", st: l === "validating_semantic_model" ? "running" : ["building", "refreshing", "saving", "verifying", "completed_live"].includes(l) ? "done" : "pending" }),
                                React.createElement(_, { n: "03", label: `Build PivotTables & visuals (${Re} created)`, st: l === "building" ? "running" : ["refreshing", "saving", "verifying", "completed_live"].includes(l) ? "done" : "pending" }),
                                React.createElement(_, { n: "04", label: "Refresh data queries", st: l === "refreshing" ? "running" : ["saving", "verifying", "completed_live"].includes(l) ? "done" : "pending" }),
                                React.createElement(_, { n: "05", label: "Save & Verify workbook", st: ["saving", "verifying"].includes(l) ? "running" : l === "completed_live" ? "done" : "pending" })
                            )
                        ),
                        l === "completed_live" && React.createElement("div", null,
                            React.createElement("div", { className: "ready-note", style: { background: "var(--teal-dim)", color: "var(--teal)", borderLeft: "3px solid var(--teal)", marginBottom: 12 } },
                                React.createElement("strong", null, "Success! Live Connection Completed."),
                                React.createElement("div", { style: { fontSize: 12, marginTop: 4 } }, "Workbook successfully upgraded to LIVE connection. Match Score: ", X !== null ? (X * 100).toFixed(0) : "100", "%.")
                            ),
                            React.createElement("div", { style: { display: "flex", gap: 8 } },
                                React.createElement("button", { className: "btn-run", onClick: Ue, style: { flex: 1, background: "var(--teal)", color: "var(--ink)" } },
                                    React.createElement(u, { id: "dl", size: 13 }), " Download Live Excel"
                                ),
                                React.createElement("button", { className: "btn-clear", onClick: qe }, "Report")
                            ),
                            React.createElement("button", { className: "btn-clear", onClick: K, style: { marginTop: 8, width: "100%", fontSize: 11 } }, "Reset Connection State")
                        ),
                        ["live_conversion_failed", "cancelled", "error"].includes(l) && React.createElement("div", null,
                            React.createElement("div", { className: "error-box", style: { marginBottom: 12 } },
                                l === "cancelled" ? "Live connection was cancelled." : `Live connection failed. State: ${l}.`,
                                ue && React.createElement("div", { style: { marginTop: 4, fontWeight: "normal", fontSize: "11px", opacity: 0.85 } }, ue)
                            ),
                            React.createElement("button", { className: "btn-run", onClick: K }, "Retry Connection")
                        )
                    )
                );
            };

            return React.createElement("div", { className: "shell", "data-mode": k }, React.createElement("header", { className: "topbar" }, React.createElement("div", { className: "logo-mark" }, React.createElement("svg", { className: "logo-svg", viewBox: "0 0 15 15", fill: "none" }, React.createElement("rect", { x: "1", y: "1", width: "5", height: "5", rx: "1", fill: "rgba(201,168,76,.6)" }), React.createElement("rect", { x: "9", y: "1", width: "5", height: "5", rx: "1", fill: "rgba(201,168,76,.4)" }), React.createElement("rect", { x: "1", y: "9", width: "5", height: "5", rx: "1", fill: "rgba(201,168,76,.4)" }), React.createElement("rect", { x: "9", y: "9", width: "5", height: "5", rx: "1", fill: "rgba(201,168,76,.8)" }))), React.createElement("div", { className: "topbar-brand" }, React.createElement("div", { className: "brand-name" }, "Conversion Studio"), React.createElement("div", { className: "brand-tagline" }, "Power BI → Excel pipeline")), React.createElement("div", { className: "topbar-divider" }), React.createElement("div", { className: "topbar-badge" }, backendStatus), React.createElement("div", { className: "topbar-spacer" }), React.createElement("div", { className: "theme-tools", "aria-label": "Appearance controls" }, React.createElement("button", { className: "mode-toggle", type: "button", onClick: () => Oe(e => e === "dark" ? "light" : "dark"), "aria-label": `Switch to ${k === "dark" ? "light" : "dark"} mode`, title: `Switch to ${k === "dark" ? "light" : "dark"} mode` }, React.createElement("span", { className: "mode-symbol", "aria-hidden": "true" }, k === "dark" ? "☀" : "☾"), React.createElement("span", { className: "mode-label" }, k === "dark" ? "Light" : "Dark"))), React.createElement("div", { className: `status-pill ${ke}` }, React.createElement("span", { className: `status-dot ${ke}` }), Je)), React.createElement("aside", { className: "left" }, React.createElement("div", { className: "left-scroll" }, React.createElement("div", { className: "card span2" }, React.createElement("div", { className: "card-head" }, React.createElement("span", { className: "card-title" }, "Workbook Summary")), React.createElement("div", { className: "card-body" }, React.createElement("div", { className: "summary-grid" }, React.createElement("div", { className: "sum-card" }, React.createElement("div", { className: "sum-n" }, g.tables), React.createElement("div", { className: "sum-lbl" }, "Tables")), React.createElement("div", { className: "sum-card" }, React.createElement("div", { className: "sum-n" }, g.relationships), React.createElement("div", { className: "sum-lbl" }, "Relations")), React.createElement("div", { className: "sum-card" }, React.createElement("div", { className: "sum-n" }, g.formulas), React.createElement("div", { className: "sum-lbl" }, "Formulas")), React.createElement("div", { className: "sum-card" }, React.createElement("div", { className: "sum-n" }, g.visuals), React.createElement("div", { className: "sum-lbl" }, "Visuals"))), React.createElement("button", { className: "btn-dl", onClick: ye, disabled: !a || w, style: { opacity: a ? 1 : .42, cursor: a ? "pointer" : "not-allowed" } }, w ? React.createElement("span", { className: "spin" }) : React.createElement(u, { id: "dl", size: 14 }), w ? "Downloading PDF…" : "Download Preview PDF"), !a && React.createElement("p", { className: "helper-note" }, "Analyze the three required files, then open Excel and connect the published Power BI semantic model."), a && React.createElement("div", { className: "ready-note" }, "Analysis complete. The Live Excel download becomes available after the live connection finishes successfully."))), React.createElement("div", { className: "card" }, React.createElement("div", { className: "card-head" }, React.createElement("span", { className: "card-title" }, "Power BI Source"), React.createElement("span", { className: "badge-req" }, "Required")), React.createElement("div", { className: "card-body" }, React.createElement("div", { className: `drop-zone ${U ? "over" : ""} ${o ? "has" : ""}`, onClick: () => { var e; return (e = $.current) == null ? void 0 : e.click() }, onDragOver: e => { e.preventDefault(), T(!0) }, onDragLeave: () => T(!1), onDrop: e => { var s; e.preventDefault(), T(!1), he((s = e.dataTransfer.files) == null ? void 0 : s[0]) } }, React.createElement("input", { ref: $, type: "file", accept: ".pbix", style: { display: "none" }, onChange: e => { var s; return he((s = e.target.files) == null ? void 0 : s[0]) } }), React.createElement("div", { className: "drop-icon" }, React.createElement(u, { id: "upload", size: 18 })), React.createElement("div", { className: "drop-label" }, o ? "File selected" : "Drop .pbix or click to browse"), React.createElement("div", { className: "drop-sub" }, o ? "Click to replace" : "Accepts .pbix only"), o && React.createElement("div", { className: "drop-filename" }, o.name)), React.createElement("button", { className: "btn-run", onClick: Xe, disabled: !o || !r || !p || L }, L && React.createElement("span", { className: "spin" }), L ? "Analyzing…" : "Analyze Files & Prepare Live Connection", !L && React.createElement(u, { id: "bolt", size: 13 })), (o || r || p) && React.createElement("button", { className: "btn-clear", onClick: He, disabled: Ne }, React.createElement(u, { id: "x", size: 12 }), " Clear all files"), de && React.createElement("div", { className: "error-box" }, de))), React.createElement("div", { className: "card" }, React.createElement("div", { className: "card-head" }, React.createElement("span", { className: "card-title" }, "Dashboard Screenshot"), React.createElement("span", { className: "badge-req" }, "Required")), React.createElement("div", { className: "card-body" }, React.createElement("div", { className: `drop-zone ${q ? "over" : ""} ${r ? "has" : ""}`, onClick: () => { var e; return (e = R.current) == null ? void 0 : e.click() }, onDragOver: e => { e.preventDefault(), ee(!0) }, onDragLeave: () => ee(!1), onDrop: e => { var s; e.preventDefault(), ee(!1), ge((s = e.dataTransfer.files) == null ? void 0 : s[0]) } }, React.createElement("input", { ref: R, type: "file", accept: ".png,.jpg,.jpeg,.webp", style: { display: "none" }, onChange: e => { var s; return ge((s = e.target.files) == null ? void 0 : s[0]) } }), z ? React.createElement("div", { className: "img-preview" }, React.createElement("img", { src: z, alt: "Preview" }), React.createElement("div", { className: "drop-filename" }, r.name), React.createElement("button", { className: "img-remove", onClick: We, type: "button" }, "✕")) : React.createElement(React.Fragment, null, React.createElement("div", { className: "drop-icon" }, React.createElement(u, { id: "img", size: 18 })), React.createElement("div", { className: "drop-label" }, "Add screenshot for theme matching"), React.createElement("div", { className: "drop-sub" }, "png · jpg · jpeg · webp"))))), React.createElement("div", { className: "card" }, React.createElement("div", { className: "card-head" }, React.createElement("span", { className: "card-title" }, "Model Metadata / TMDL"), React.createElement("span", { className: "badge-req" }, "Required")), React.createElement("div", { className: "card-body" }, React.createElement("div", { className: `drop-zone ${$e ? "over" : ""} ${p ? "has" : ""}`, onClick: () => { var e; return (e = x.current) == null ? void 0 : e.click() }, onDragOver: e => { e.preventDefault(), se(!0) }, onDragLeave: () => se(!1), onDrop: e => { var s; e.preventDefault(), se(!1), we((s = e.dataTransfer.files) == null ? void 0 : s[0]) } }, React.createElement("input", { ref: x, type: "file", accept: ".tmdl,.txt,.json", style: { display: "none" }, onChange: e => { var s; return we((s = e.target.files) == null ? void 0 : s[0]) } }), React.createElement("div", { className: "drop-icon" }, React.createElement(u, { id: "upload", size: 18 })), React.createElement("div", { className: "drop-label" }, p ? "TMDL metadata selected" : "Add TMDL/model metadata"), React.createElement("div", { className: "drop-sub" }, p ? "Click to replace" : "tmdl · txt · json"), p && React.createElement("div", { className: "drop-filename", style: { marginTop: 10, position: "relative", paddingRight: 26 } }, p.name, React.createElement("button", { className: "img-remove", onClick: Ve, type: "button", style: { top: -6, right: -6 } }, "✕"))))), React.createElement("div", { className: "card" }, React.createElement("div", { className: "card-head" }, React.createElement("span", { className: "card-title" }, "Pipeline")), React.createElement("div", { className: "card-body" }, React.createElement("div", { className: "steps" }, React.createElement(_, { n: "01", label: "Read PBIX, screenshot & TMDL", st: V.a }), React.createElement(_, { n: "02", label: "Analyze visuals & model fields", st: V.b }), React.createElement(_, { n: "03", label: "Prepare preview & live session", st: V.c }), React.createElement(_, { n: "04", label: "Ready to open Excel", st: V.d })))), renderLiveConnectCard())), React.createElement("main", { className: "right" }, React.createElement("div", { className: "viewer-head" }, React.createElement("div", { className: "viewer-title-wrap" }, React.createElement("div", { className: "workspace-kicker" }, "Conversion workspace"), React.createElement("div", { className: "viewer-title" }, "Report Analysis & Preview"), React.createElement("div", { className: "viewer-sub" }, a ? "Conversion complete — browse converted chunks below" : "Waiting for PBIX upload and processing")), React.createElement("button", { className: "btn-dl-top", onClick: ye, disabled: !a || w, style: { opacity: a ? 1 : .4, cursor: a ? "pointer" : "not-allowed" } }, w ? React.createElement("span", { className: "spin" }) : React.createElement(u, { id: "dl", size: 13 }), w ? "PDF…" : "Download PDF")), React.createElement("nav", { className: "tab-bar" }, Ge.map(e => React.createElement("button", { key: e.id, className: `tab ${ce === e.id ? "active" : ""}`, onClick: () => W(e.id) }, e.label, e.n !== null && React.createElement("span", { className: "tab-cnt" }, e.n)))), React.createElement("div", { className: "content" }, React.createElement(as, { data: a, tab: ce })), React.createElement("footer", { className: "app-footer" }, React.createElement("div", { className: "footer-left" }, React.createElement("span", { className: "footer-dot" }), React.createElement("span", null, `Conversion Studio · Power BI to Excel workspace (${backendStatus})`)), React.createElement("div", { className: "footer-right" }, React.createElement("span", null, a ? `Session ${a.session_id || "ready"}` : "No active session")))))
        }; ReactDOM.createRoot(document.getElementById("root")).render(React.createElement(ts, null));
})();
