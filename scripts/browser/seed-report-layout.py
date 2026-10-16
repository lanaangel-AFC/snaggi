# WP8: seeds a temporary new-version project with two inspections, so the new report layout
# has records that are new, carried in, amended, closed (new and carried), a note, an
# uncategorised observation, a category left out of the Action List, a project instruction,
# a narrative and three elevation drawings (two images and one PDF, plus one with no pins).
# Also seeds an old-version project with two inspections. Writes /tmp/wp8seed.json.
# Usage: python3 scripts/browser/seed-report-layout.py [BASE]
import json, urllib.request, io, uuid, sys
from PIL import Image, ImageDraw

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:5070"
C = "2026-10-15T00:00:00.000Z"

def call(m, p, b=None):
    r = urllib.request.Request(B + p, data=json.dumps(b).encode() if b is not None else None, headers={"Content-Type": "application/json"}, method=m)
    raw = urllib.request.urlopen(r, timeout=60).read()
    return json.loads(raw) if raw else None

def upload(p, field, fname, data, mime, extra={}):
    bd = uuid.uuid4().hex; body = b""
    for k, v in extra.items():
        body += f"--{bd}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode()
    body += f"--{bd}\r\nContent-Disposition: form-data; name=\"{field}\"; filename=\"{fname}\"\r\nContent-Type: {mime}\r\n\r\n".encode() + data + f"\r\n--{bd}--\r\n".encode()
    r = urllib.request.Request(B + p, data=body, headers={"Content-Type": f"multipart/form-data; boundary={bd}"}, method="POST")
    return json.load(urllib.request.urlopen(r, timeout=60))

def img(fmt, color, size, text=""):
    b = io.BytesIO(); im = Image.new("RGB", size, color); d = ImageDraw.Draw(im)
    d.rectangle([size[0] // 8, size[1] // 8, size[0] * 7 // 8, size[1] * 7 // 8], outline=(0, 0, 0), width=6)
    for i in range(1, 6):
        d.line([size[0] // 8, size[1] * i // 6, size[0] * 7 // 8, size[1] * i // 6], fill=(90, 90, 90), width=3)
    if text: d.text((size[0] // 8 + 10, size[1] // 8 + 10), text, fill=(0, 0, 0))
    im.save(b, fmt); return b.getvalue()

def record(pid, rid, uid, **kw):
    e, dr, lv, wt, sq = uid.split("-")
    body = {"reportId": rid, "dateOpened": "2026-10-15", "verificationMethod": "Visual", "verificationPerson": "AFC", "status": "open",
            "createdAt": C, "elevationCode": e, "dropCode": dr, "levelCode": lv, "workTypeCode": wt, "seqNumber": sq,
            "uidPrefix": f"{e}-{dr}-{lv}-{wt}", "uidOverride": uid, "recordType": "defect", "assignedTo": "Contractor", "dueDate": "2026-10-23"}
    body.update(kw)
    return call("POST", f"/api/projects/{pid}/defects", body)

def photo(did, slot, color):
    upload(f"/api/defects/{did}/photos", "photo", f"{slot}.jpg", img("JPEG", color, (800, 600)), "image/jpeg", {"slot": slot, "caption": f"Caption {slot}"})

def project(name, fv):
    body = {"name": name, "address": "1 Layout St", "client": "Test Client", "inspector": "LA", "createdAt": C, "afcReference": "AFC-26002"}
    if fv: body["featureVersion"] = fv
    pid = call("POST", "/api/projects", body)["id"]
    call("PUT", f"/api/projects/{pid}/categories", [
        {"code": "RR", "label": "Rectify", "isDefault": True},
        {"code": "RD", "label": "Redesign"},
        {"code": "RFI", "label": "Request Information", "inActionList": True},
        {"code": "WIP", "label": "Work in progress"},
        {"code": "MS", "label": "Monitor", "inActionList": False},
        {"code": "UN", "label": "Unused category"},
        {"code": "PI", "label": "Project Instruction"},
    ])
    r1 = call("POST", f"/api/projects/{pid}/reports", {"inspectionNumber": "01", "inspectionDate": "2026-10-08", "createdAt": C, "locationsCovered": "North Elevation, Drop 1"})["id"]
    a = record(pid, r1, "N-01-05-CR-01", comment="Spalled concrete at slab edge", actionRequired="Break out and patch repair", categoryCode="RR")
    photo(a["id"], "wip1", (200, 120, 120)); photo(a["id"], "wip2", (180, 140, 120))
    b = record(pid, r1, "N-01-04-CR-02", comment="Exposed reinforcement", actionRequired="Treat and patch", categoryCode="RR")
    c = record(pid, r1, "S-02-05-GL-01", comment="Glazing bead detail does not suit the frame", actionRequired="Engineer to redesign bead detail", categoryCode="RD")
    record(pid, r1, "E-01-01-WP-01", comment="Membrane in progress", actionRequired="Continue", categoryCode="WIP")
    record(pid, r1, "N-01-02-CR-03", comment="Crack at window head", actionRequired="Rake out and seal", categoryCode="RR")
    record(pid, r1, "N-01-03-SE-01", recordType="observation", comment="Sealant type not confirmed", actionRequired="Contractor to provide sealant data sheet", categoryCode="RFI")
    r2 = call("POST", f"/api/reports/{r1}/start-next-inspection", {"inspectionNumber": "02", "inspectionDate": "2026-10-15", "locationsCovered": "North Elevation"})["id"]
    rows = {d["uid"]: d for d in call("GET", f"/api/reports/{r2}/defects")}
    call("PATCH", f"/api/defects/{rows['N-01-02-CR-03']['id']}", {"status": "complete", "dateClosed": "2026-10-15", "contextReportId": r2})
    call("PATCH", f"/api/defects/{rows['N-01-04-CR-02']['id']}", {"comment": "Exposed reinforcement, now wider", "contextReportId": r2})
    n1 = record(pid, r2, "N-02-05-CR-04", comment="New spall below balcony", actionRequired="Patch repair", categoryCode="RR", dateOpened="2026-10-15")
    for s, col in (("wip1", (120, 160, 200)), ("wip2", (120, 200, 160)), ("wip3", (200, 200, 120))): photo(n1["id"], s, col)
    n2 = record(pid, r2, "S-03-02-GL-02", comment="Glass clamp spacing exceeds detail", actionRequired="Engineer to review clamp layout", categoryCode="RD", dateOpened="2026-10-15")
    n3 = record(pid, r2, "N-02-01-SE-02", recordType="observation", comment="Sealant joints wider than 25 mm", actionRequired="Confirm joint design", dateOpened="2026-10-15")
    n4 = record(pid, r2, "E-02-02-GW-01", comment="Hairline crack to monitor", actionRequired="Monitor at next visit", categoryCode="MS", dateOpened="2026-10-15")
    n5 = record(pid, r2, "E-03-03-GW-02", comment="Loose fixing found and fixed", actionRequired="Refix", categoryCode="RR", dateOpened="2026-10-15")
    call("PATCH", f"/api/defects/{n5['id']}", {"status": "complete", "dateClosed": "2026-10-15", "contextReportId": r2})
    if fv != 1:  # notes exist in new projects only
        note = record(pid, r2, "N-02-02-GW-03", recordType="observation", actionNeeded=False, comment="Substrate preparation not carried into corners", actionRequired="", dateOpened="2026-10-15")
        photo(note["id"], "wip1", (150, 150, 150))
    els = {}
    els["N"] = upload(f"/api/projects/{pid}/elevations", "file", "n.png", img("PNG", (245, 245, 245), (1600, 1000), "NORTH"), "image/png", {"name": "North elevation"})["id"]
    els["S"] = upload(f"/api/projects/{pid}/elevations", "file", "s.png", img("PNG", (235, 240, 250), (1000, 1400), "SOUTH"), "image/png", {"name": "South elevation"})["id"]
    els["E"] = upload(f"/api/projects/{pid}/elevations", "file", "e.pdf", img("PDF", (250, 250, 240), (1400, 900), "EAST"), "application/pdf", {"name": "East elevation (PDF)"})["id"]
    els["W"] = upload(f"/api/projects/{pid}/elevations", "file", "w.png", img("PNG", (250, 250, 250), (800, 600), "WEST"), "image/png", {"name": "West elevation (no pins)"})["id"]
    rows = {d["uid"]: d for d in call("GET", f"/api/reports/{r2}/defects")}
    for d in list(rows.values()):
        if d["uid"].startswith("Note-"): rows["N-02-02-GW-03"] = d  # the server names notes itself
    pins = [("N", "N-01-05-CR-01", 20, 30), ("N", "N-01-04-CR-02", 40, 50), ("N", "N-02-05-CR-04", 60, 30), ("N", "N-01-02-CR-03", 75, 70), ("N", "N-02-02-GW-03", 30, 75),
            ("S", "S-02-05-GL-01", 40, 40), ("S", "S-03-02-GL-02", 60, 60), ("E", "E-02-02-GW-01", 50, 50), ("E", "E-03-03-GW-02", 30, 40)]
    for el, uid, x, yv in pins:
        if uid in rows:
            call("POST", f"/api/elevations/{els[el]}/markers", {"defectId": rows[uid]["id"], "defectUid": rows[uid]["uid"], "status": "open", "xPercent": x, "yPercent": yv})
    if fv != 1:
        call("POST", f"/api/projects/{pid}/instructions", {"reportId": r2, "workTypes": ["CR"], "observation": "Patch repairs are being finished proud of the surface", "actionRequired": "Finish all patch repairs flush with the surrounding concrete", "responsibility": "Contractor", "dateIssued": "2026-10-15"})
    call("POST", f"/api/projects/{pid}/narrative-holds", {"title": "Scaffold access", "body": "Scaffold to the north elevation is due to be dismantled on 30 October.", "status": "Active", "dateRaised": "2026-10-15"})
    try: call("POST", f"/api/reports/{r2}/action-summaries/regenerate")
    except Exception as e: print("regenerate skipped", e)
    return pid, r1, r2, els

P, R1, R2, ELS = project("ZZ Report layout check (temporary)", None)
OP, OR1, OR2, _ = project("ZZ Old layout check (temporary)", 1)
out = dict(P=P, R1=R1, R2=R2, ELS=ELS, OP=OP, OR1=OR1, OR2=OR2)
json.dump(out, open("/tmp/wp8seed.json", "w")); print(out)
