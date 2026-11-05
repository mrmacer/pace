///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 21, 2026
// Filename: diagnostic.js
// Purpose: Temporary, read-only production diagnostic for validating SharePoint column
//          names, types, and populated/blank status against the PACE Visits list. It
//          deliberately reports metadata only and must never include tokens, student
//          names, IDs, emails, or field contents, and must be removed once the live
//          schema is fully confirmed.
///////////////////////////////////////////////////////////////////////////////////////////////

// Read-only: uses only existing GRAPH/AUTH infrastructure (no new Graph calls beyond a
// plain column-list GET, which graph.js doesn't already expose a helper for). Never
// creates, updates, or deletes anything, never displays a token, and never displays a
// student name/ID/email or any actual field VALUE -- only column metadata and
// POPULATED/BLANK status.
//
// To remove this feature entirely once field mapping is reconciled: delete this file,
// its <script> tag in index.html, the "Run SharePoint Diagnostic" button + overlay
// markup in index.html, and the three diagnostic event-listener blocks in app.js (all
// marked "PATCH 003 DIAGNOSTIC").

// ROOM-FIELD-NAME PATCH: "Room" added — the user manually confirmed this
// is the live column's actual display name. "PACE Room"/"Pace Room" stay
// in this list too so a re-run of this diagnostic keeps showing (rather
// than silently hiding) whichever of the three names turns out NOT to be
// live, exactly like the "SCM Used"/"SCM" pair already does below.
const DIAGNOSTIC_CONCEPTS = [
  "Student", "Student ID", "Room", "PACE Room", "Pace Room", "Date", "Time In", "Time Out", "Duration",
  "Staff Member", "Behavior Specialist", "Reason", "Behavior", "Interventions",
  "Support", "SCM Used", "SCM", "Notes", "Entry ID", "Submitted By", "Submitted At"
];

const Diagnostic = {
  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: run
  // Description: Fetches only the live SharePoint list schema and reduces it to metadata
  //   only (column names/types) for the diagnostic report. It deliberately does not read
  //   any visit rows, student records, or note contents.
  // Parameters: none
  ///////////////////////////////////////////////////////////////////////////////////////////////
  async run() {
    // Defense in depth: GRAPH._get() already throws in demo mode via
    // assertGraphAllowed(), but fail fast and explicitly here too.
    if (APP_MODE === "demo") throw new Error("Diagnostic is production-only.");

    const siteId = await GRAPH.getSiteId();
    const listId = await GRAPH.getListId(CONFIG.LISTS.paceVisits);
    const token  = await AUTH.acquireGraphToken();

    // graph.js doesn't have a "columns with type" helper (getListSchema()
    // only pulls name/displayName) — this is the one new read-only GET,
    // same endpoint pattern as everything else in graph.js.
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    const timeout = setTimeout(() => controller?.abort(), 15000);
    let resp;
    try {
      resp = await fetch(
        `${GRAPH._BASE}/sites/${siteId}/lists/${listId}/columns`,
        {
          headers: { Authorization: `Bearer ${token}` },
          ...(controller ? { signal: controller.signal } : {})
        }
      );
    } catch (err) {
      if (err?.name === "AbortError") throw new Error("Column metadata request timed out.");
      throw err;
    } finally {
      clearTimeout(timeout);
    }
    if (!resp.ok) throw new Error(`Column fetch failed: ${resp.status}`);
    const data = await resp.json();

    // PATCH 010: a bare "text" type doesn't say whether a column is
    // SharePoint's "Single line of text" (hard 255-character cap) or
    // "Multiple lines of text" (no such cap) — Graph represents both as a
    // `text` facet, distinguished only by `text.allowMultipleLines`/
    // `text.maxLength`. This was the missing piece needed to answer the
    // Notes-length question conclusively instead of guessing — see
    // README "Notes column type — still unconfirmed."
    const columns = (data.value || []).map(c => ({
      displayName: c.displayName,
      internalName: c.name,
      type: c.text ? "text" : c.choice ? "choice" : c.dateTime ? "dateTime"
          : c.boolean ? "boolean" : c.number ? "number" : c.personOrGroup ? "personOrGroup"
          : c.lookup ? "lookup" : "other",
      textDetail: c.text
        ? `allowMultipleLines: ${c.text.allowMultipleLines === true}, maxLength: ${c.text.maxLength ?? "(default 255)"}`
        : null
    }));

    // Does a column with this EXACT display name exist? (Never guesses —
    // just reports what's actually there.)
    const conceptMap = DIAGNOSTIC_CONCEPTS.map(concept => {
      const match = columns.find(c => c.displayName === concept);
      return match
        ? { concept, status: "FOUND", internalName: match.internalName, type: match.type, textDetail: match.textDetail }
        : { concept, status: "NOT FOUND" };
    });

    return { columns, conceptMap, latestFields: null, visitCount: null };
  },

  ///////////////////////////////////////////////////////////////////////////////////////////////
  // Function Name: formatReport
  // Description: Formats the schema and latest-visit metadata produced by run() into the
  //   copyable plain-text diagnostic report.
  // Parameters: object result - the diagnostic result returned by run(), containing
  //   columns, conceptMap, latestFields, and visitCount - input
  ///////////////////////////////////////////////////////////////////////////////////////////////
  formatReport(result) {
    const lines = [];
    lines.push("TEMPORARY PATCH 003 DIAGNOSTIC — IEP_Pace_Visits");
    lines.push(`Generated: ${new Date().toISOString()}`);
    lines.push("No student names, IDs, emails, tokens, or record values are included below.");
    lines.push("");
    lines.push("=== A. Live column schema ===");
    result.columns.forEach(c => {
      lines.push(`${c.displayName}  |  internal: ${c.internalName}  |  type: ${c.type}${c.textDetail ? `  |  ${c.textDetail}` : ""}`);
    });
    lines.push("");
    lines.push("=== B. Requested concept -> live column lookup ===");
    result.conceptMap.forEach(c => {
      lines.push(c.status === "FOUND"
        ? `${c.concept}: FOUND — internal "${c.internalName}", type ${c.type}${c.textDetail ? ` (${c.textDetail})` : ""}`
        : `${c.concept}: NOT FOUND`);
    });
    lines.push("");
    lines.push("=== C. Visit data not fetched ===");
    lines.push("This metadata-only diagnostic does not read student records or note contents.");
    return lines.join("\n");
  }
};
