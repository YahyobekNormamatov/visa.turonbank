import re
import pdfplumber

NOISE = [
    re.compile(r"^Card Suite$"),
    re.compile(r"^Clearing$"),
    re.compile(r"^VSS - Electronic Settlement Detail Report"),
    re.compile(r"^System Processing Date:"),
    re.compile(r"^Report Date:"),
    re.compile(r"^Created \d{4}-\d{2}-\d{2}"),
    re.compile(r"^VISA-R_IN_VSS_REP_DATA_TURON"),
]
NUM_RE = re.compile(r'^-?[\d,]+\.\d{2}$')


def is_noise(line):
    return any(p.match(line) for p in NOISE)


def to_float(tok):
    s = tok.strip().replace(",", "")
    if not s or s in ("-", "—"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def smart_tokens(line):
    parts = line.split()
    nums = []
    i = len(parts)
    while i > 0 and NUM_RE.match(parts[i - 1]):
        nums.insert(0, parts[i - 1])
        i -= 1
    label = " ".join(parts[:i])
    return label, nums


def extract_text_from_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text


def clean_lines(raw_text):
    lines = []
    for raw in raw_text.split("\n"):
        line = raw.strip()
        if line and not is_noise(line):
            lines.append(line)
    return lines


def parse(raw_text):
    lines = clean_lines(raw_text)
    rows = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]

        if line in ("INTERNATIONAL SETTLEMENT", "NATIONAL SETTLEMENT"):
            rows.append({"t": "banner", "text": line})
            i += 1; continue

        m = re.match(r"^Reporting for:\s*(.+?)\s+Settlement Currency Code:\s*(\S+)$", line)
        if m:
            rows.append({"t": "kv2", "k1": "Reporting for", "v1": m.group(1),
                         "k2": "Currency", "v2": m.group(2)})
            i += 1; continue

        m = re.match(r"^Settl\. Date:\s*(\S+)\s+Clearing Currency Code:\s*(\S+)$", line)
        if m:
            rows.append({"t": "kv2", "k1": "Settl. Date", "v1": m.group(1),
                         "k2": "Clearing Currency", "v2": m.group(2)})
            i += 1; continue

        m = re.match(r"^Settl\. Date:\s*(\S+)$", line)
        if m:
            rows.append({"t": "kv2", "k1": "Settl. Date", "v1": m.group(1),
                         "k2": "", "v2": ""})
            i += 1; continue

        if line.endswith("Count Interch. value Reimb. Fees Net value"):
            sec = line[: -len("Count Interch. value Reimb. Fees Net value")].strip()
            rows.append({"t": "h4", "sec": sec,
                         "cols": ["Count", "Interch. value", "Reimb. Fees", "Net value"]})
            i += 1; continue

        if line == "Optional Issuer conversion fees Interch. amount Convers. Fee Opt. issuer fee":
            rows.append({"t": "h3opt", "sec": "Optional Issuer conversion fees",
                         "cols": ["Interch. amount", "Convers. Fee", "Opt. issuer fee"]})
            i += 1; continue

        if line in ("Visa Charges (Issuer) Amount", "Visa Charges (Acquirer) Amount"):
            rows.append({"t": "h1", "sec": line.rsplit(" ", 1)[0], "cols": ["Amount"]})
            i += 1; continue

        if line.startswith("ISA CHARGE") or line.startswith("IAF CHARGE"):
            label, nums = smart_tokens(line)
            rows.append({"t": "d1", "label": label,
                         "val": to_float(nums[0]) if nums else None})
            i += 1; continue

        if line.startswith("Total for VISA charges"):
            _, nums = smart_tokens(line)
            rows.append({"t": "tot1", "label": "Total for VISA charges",
                         "val": to_float(nums[0]) if nums else None})
            i += 1; continue

        if line.startswith("Net Settlement Amount"):
            _, nums = smart_tokens(line)
            rows.append({"t": "tot1", "label": line.split("(")[0].strip() + " (" + line.split("(")[1] if "(" in line else line,
                         "val": to_float(nums[0]) if nums else None})
            i += 1; continue

        if line.startswith("Funds Transfer Amount"):
            _, nums = smart_tokens(line)
            rows.append({"t": "tot1", "label": "Funds Transfer Amount",
                         "val": to_float(nums[0]) if nums else None})
            i += 1; continue

        if line.startswith("Total:"):
            label, nums = smart_tokens(line)
            vals = [to_float(x) for x in nums]
            while len(vals) < 4:
                vals.insert(0, None)
            cnt = vals[-4]
            m2 = re.match(r"^Total:\s+(\d+)$", label)
            if m2 and cnt is None:
                cnt = float(m2.group(1))
            rows.append({"t": "tot4", "label": "Total:",
                         "count": cnt, "interch": vals[-3],
                         "reimb": vals[-2], "net": vals[-1]})
            i += 1; continue

        if line.startswith("Total ") and not line.startswith("Total:"):
            label, nums = smart_tokens(line)
            if label == "Total" and len(nums) == 3:
                vals = [to_float(x) for x in nums]
                rows.append({"t": "tot3", "label": "Total",
                             "interch": vals[0], "conv": vals[1], "opt": vals[2]})
                i += 1; continue

        if line.startswith("Originals "):
            _, nums = smart_tokens(line)
            if len(nums) == 3:
                vals = [to_float(x) for x in nums]
                rows.append({"t": "d3", "label": "Originals",
                             "interch": vals[0], "conv": vals[1], "opt": vals[2]})
                i += 1; continue

        if re.match(r"^\d+\s+-?[\d,]+\.\d{2}\s+-?[\d,]+\.\d{2}\s+-?[\d,]+\.\d{2}$", line):
            parts = line.split()
            rows.append({"t": "tot4plain",
                         "count": to_float(parts[0]), "interch": to_float(parts[1]),
                         "reimb": to_float(parts[2]), "net": to_float(parts[3])})
            i += 1; continue

        label, nums = smart_tokens(line)
        if nums:
            vals = [to_float(x) for x in nums]
            while len(vals) < 4:
                vals.insert(0, None)
            cnt, interch, reimb, net = vals[-4], vals[-3], vals[-2], vals[-1]
            m2 = re.match(r"^(.*\S)\s+(\d+)$", label)
            if m2 and cnt is None:
                label = m2.group(1)
                cnt = float(m2.group(2))
            rows.append({"t": "d4", "label": label,
                         "count": cnt, "interch": interch, "reimb": reimb, "net": net})
        else:
            rows.append({"t": "text", "label": line})
        i += 1

    return rows
