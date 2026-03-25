import openpyxl
from openpyxl import Workbook
import random
from collections import Counter

# ── User inputs ────────────────────────────────────────────────────────────────

while True:
    try:
        total = int(input("How many contractors do you want to generate? "))
        if total < 1:
            print("  Please enter a number greater than 0.")
            continue
        break
    except ValueError:
        print("  Invalid input. Please enter a whole number.")

print()
job_title_input = input("Job Title (required - e.g. Carpenter, Foreman, Journeyman Carpenter): ").strip()
while not job_title_input:
    job_title_input = input("  Job Title is required: ").strip()

print()
union_input = input("Union Name (press Enter to leave blank): ").strip()

print()
job_level_input = input("Job Level (press Enter to leave blank): ").strip()

print()
classification_input = input("Job Classification (press Enter to leave blank): ").strip()

union_filled = bool(union_input or job_level_input or classification_input)

def ask_pay_rates():
    print()
    print("  Hourly rates (used for 90% weekly contractors):")
    while True:
        try:
            hr = float(input("    Compensation Rate: "))
            break
        except ValueError:
            print("    Invalid. Enter a number.")
    while True:
        try:
            ot = float(input("    Overtime Rate: "))
            break
        except ValueError:
            print("    Invalid. Enter a number.")
    print("  Annual rates (used for 10% annually contractors):")
    while True:
        try:
            ar = float(input("    Compensation Rate: "))
            break
        except ValueError:
            print("    Invalid. Enter a number.")
    while True:
        try:
            ao = float(input("    Overtime Rate: "))
            break
        except ValueError:
            print("    Invalid. Enter a number.")
    return hr, ot, ar, ao

print()
if union_filled:
    print("Pay rates (optional — press Enter to leave blank):")
    print("  A. Enter pay rates manually")
    print("  B. Leave pay rate columns blank")
    while True:
        pay_choice = input("Select pay rate option (A/B): ").strip().upper()
        if pay_choice in ("A", "B"):
            break
        print("  Please enter A or B.")
    if pay_choice == "A":
        hourly_rate, hourly_ot, annual_rate, annual_ot = ask_pay_rates()
    else:
        hourly_rate = hourly_ot = annual_rate = annual_ot = ""
else:
    print("Union / Job Level / Classification are blank — pay rates are required.")
    print("  A. Enter pay rates manually")
    print("  B. Use default rates (hourly: $25.00 / OT: $37.50 | annually: $60,000 / OT: $0)")
    while True:
        pay_choice = input("Select pay rate option (A/B): ").strip().upper()
        if pay_choice in ("A", "B"):
            break
        print("  Please enter A or B.")
    if pay_choice == "A":
        hourly_rate, hourly_ot, annual_rate, annual_ot = ask_pay_rates()
    else:
        hourly_rate = 25.0
        hourly_ot   = 37.5
        annual_rate = 60000.0
        annual_ot   = 0.0

# ── Static data ────────────────────────────────────────────────────────────────

random.seed(42)

ADDRESSES = [
    "691 Green Crest Drive, Westerville, OH, US, 43081",
    "479 Clifden Ct, Sunbury, OH, US, 43074",
    "480 Clifden Ct, Sunbury, OH, US, 43074",
    "481 Clifden Ct, Sunbury, OH, US, 43074",
    "2065 Jervis Road, Columbus, OH, US, 43221",
    "2066 Jervis Road, Columbus, OH, US, 43221",
    "2067 Jervis Road, Columbus, OH, US, 43221",
    "919 West Mulberry Street, Lancaster, OH, US, 43130",
    "342 Oak Hill Drive, Pickerington, OH, US, 43147",
    "1204 Maple Street, Grove City, OH, US, 43123",
    "88 Shady Lane, Heath, OH, US, 43056",
    "657 Elm Court, Newark, OH, US, 43055",
    "234 Riverside Drive, Chillicothe, OH, US, 45601",
    "510 Fairview Avenue, Delaware, OH, US, 43015",
    "1801 Pine Ridge Road, Zanesville, OH, US, 43701",
    "407 Briarwood Drive, Circleville, OH, US, 43113",
    "733 Chestnut Street, Ashville, OH, US, 43103",
    "156 Hilltop Drive, Johnstown, OH, US, 43031",
    "922 Walnut Street, Pataskala, OH, US, 43062",
    "314 Cedar Run, Mount Vernon, OH, US, 43050",
]

male_first = [
    "Travis","Randy","Cody","Shane","Tyler","Bobby","Derek","Dale","Jimmy","Mike",
    "Dave","Steve","Rick","Tony","Joe","Danny","Chris","Scott","Brian","Chad",
    "Brad","Kyle","Justin","Brandon","Blake","Gary","Ronnie","Dustin","Kevin","Keith",
    "Todd","Craig","Brett","Lance","Kurt","Wade","Troy","Brent","Greg","Kirk",
    "Dean","Ross","Drew","Kent","Clint","Vince","Gene","Earl","Floyd","Glen",
    "Lester","Merle","Darrell","Lonnie","Dwight","Marvin","Harvey","Howard","Norman","Roy",
    "Ray","Rex","Russ","Zach","Zane","Luke","Lyle","Lee","Jay","Joel",
    "Josh","Jack","Jake","Jared","Jason","Jeff","Jerry","Jesse","Jim","John",
    "Jon","Jordan","Julian","Karl","Ken","Larry","Leon","Lewis","Lloyd",
    "Logan","Mark","Mason","Matt","Max","Mitch","Morgan","Nathan","Neil","Nick",
    "Nolan","Owen","Paul","Perry","Pete","Phil","Preston","Ralph","Reed","Reid",
    "Ricky","Rob","Rod","Roger","Roland","Ron","Rory","Rudy","Russell","Ryan",
    "Sam","Seth","Shawn","Skip","Skyler","Spencer","Stan","Tanner","Taylor","Ted",
    "Terry","Thomas","Tim","Tom","Trent","Trevor","Tucker","Ty","Wayne","Wes",
]

female_first = [
    "Crystal","Tammy","Brenda","Ashley","Kayla","Jessica","Brittany","Amber","Heather","Melissa",
    "Tiffany","Nicole","Rachel","Sarah","Lauren","Megan","Amanda","Stephanie","Holly","Courtney",
    "Becky","Cindy","Debbie","Donna","Karen","Linda","Lisa","Lori","Lynn","Mary",
    "Nancy","Pam","Patricia","Sandy","Sharon","Sherri","Susan","Tara","Teresa","Valerie",
    "Whitney","Alicia","Angela","April","Barbara","Betty","Bonnie","Carol","Carrie","Denise",
]

last_names = [
    "Briggs","Harmon","Norris","Crawford","Mercer","Tanner","Sutton","Garrett","Fowler","Webb",
    "Carter","Walker","Hayes","Coleman","Griffin","Simmons","Ford","Holt","Kerr","Doyle",
    "Pratt","Bishop","Hughes","Payne","Anderson","Baker","Bennett","Brooks","Brown","Burns",
    "Butler","Campbell","Chapman","Clark","Collins","Cook","Cooper","Cox","Curtis","Davis",
    "Dixon","Edwards","Ellis","Evans","Ferguson","Fisher","Fleming","Fletcher","Foster","Fox",
    "Freeman","Fuller","Gibson","Graham","Grant","Gray","Green","Hall","Hamilton","Hansen",
    "Harris","Hart","Harvey","Henderson","Hill","Holmes","Howard","Hudson","Hunt","Hunter",
    "James","Jensen","Johnson","Jones","Jordan","Kelly","Kennedy","King","Lambert","Lane",
    "Lawrence","Lewis","Long","Lynch","Martin","Mason","Matthews","Maxwell","Miller","Mitchell",
    "Moore","Morgan","Morris","Morrison","Murphy","Murray","Myers","Nelson","Newton","Nichols",
    "Owens","Parker","Patterson","Pearson","Peters","Phillips","Pierce","Porter","Powell","Price",
    "Quinn","Reed","Reynolds","Richards","Richardson","Riley","Roberts","Robertson","Robinson","Rogers",
    "Ross","Russell","Sanders","Scott","Shaw","Sherman","Simpson","Smith","Spencer","Stevens",
    "Stewart","Stone","Sullivan","Taylor","Thomas","Thompson","Turner","Wagner","Warren","Watson",
    "Wells","White","Williams","Wilson","Wood","Wright","Young","Zimmerman","Adkins","Allen",
    "Arnold","Austin","Barker","Barnes","Barrett","Bates","Bauer","Bell","Berry","Black",
    "Blair","Blake","Bowen","Boyd","Bradford","Bradley","Brady","Brock","Burke","Carr",
    "Carroll","Chandler","Chase","Church","Cobb","Combs","Conrad","Conner","Dalton","Dawson",
    "Douglas","Drake","Dunn","Eaton","Farrell","Fields","Figueroa","Finley","Flynn","Frost",
]


# ── Uniqueness tracking ────────────────────────────────────────────────────────

used_names  = set()
used_codes  = set()
used_phones = set()
used_emails = set()
used_ssn    = set()

def gen_code():
    while True:
        digits = random.randint(4, 5)
        c = str(random.randint(10**(digits-1), 10**digits - 1))
        if c not in used_codes:
            used_codes.add(c)
            return c

def gen_phone():
    while True:
        area = random.choice(["614","740","513","330","216","419","937","234","380","567"])
        num = "+1" + area + str(random.randint(1000000, 9999999)).zfill(7)
        if num not in used_phones:
            used_phones.add(num)
            return num

def gen_email(fn, ln):
    suffixes = ["","01","02","03","_oh","_acp","_wrk","_pro","_2026","_x"]
    domains  = ["gmail.com","yahoo.com","outlook.com","hotmail.com"]
    base = fn.lower() + "." + ln.lower().replace(" ", "")
    for sfx in suffixes:
        for dom in domains:
            email = base + sfx + "@" + dom
            if email not in used_emails:
                used_emails.add(email)
                return email
    while True:
        email = base + str(random.randint(100, 999)) + "@gmail.com"
        if email not in used_emails:
            used_emails.add(email)
            return email

def gen_dob():
    y = random.randint(1968, 2000)
    m = random.randint(1, 12)
    d = random.randint(1, 28)
    return f"{y}-{m:02d}-{d:02d}"

def gen_ssn():
    while True:
        s = "XXX-XX-" + str(random.randint(1000, 9999))
        if s not in used_ssn:
            used_ssn.add(s)
            return s

def gen_name(gender):
    pool = male_first if gender == "MALE" else female_first
    for _ in range(5000):
        fn = random.choice(pool)
        ln = random.choice(last_names)
        full = fn + " " + ln
        if full not in used_names:
            used_names.add(full)
            return fn, ln
    raise Exception("Name pool exhausted")

# ── Build gender list (85% male / 15% female) ─────────────────────────────────

male_count   = max(1, round(total * 0.85))
female_count = total - male_count
genders = ["MALE"] * male_count + ["FEMALE"] * female_count
random.shuffle(genders)

# ── Build pay-frequency list (90% Weekly / 10% Annually) ─────────────────────

annually_count = max(1, round(total * 0.10))
weekly_count   = total - annually_count
pay_freqs = ["Weekly"] * weekly_count + ["Annually"] * annually_count
random.shuffle(pay_freqs)

# ── Generate contractors ───────────────────────────────────────────────────────

contractors = []
for idx, (gender, pay_freq) in enumerate(zip(genders, pay_freqs)):
    fn, ln  = gen_name(gender)
    code    = gen_code()
    phone   = gen_phone()
    email   = gen_email(fn, ln)
    dob     = gen_dob()
    ssn     = gen_ssn()
    address = random.choice(ADDRESSES)

    hire_date = "2026-03-" + str(random.randint(1, 20)).zfill(2)

    # Pay rates
    if pay_freq == "Annually":
        comp_type = "annually"
        comp_rate = annual_rate
        ot_rate   = annual_ot
    else:
        comp_type = "hourly"
        comp_rate = hourly_rate
        ot_rate   = hourly_ot

    # 38 data columns + 4 empty = 42 total (matching template)
    row_data = [
        fn, "", ln, job_title_input, code, gender, dob, hire_date, hire_date, "",
        "Worker", phone, email, address,
        "Contractor", True, False, True,
        comp_type, pay_freq, str(comp_rate), str(ot_rate),
        "", "",                    # Lumber Id, External Id
        "", "",                    # Branch Code, Branch Name
        union_input, job_level_input, classification_input,
        "", "", "", "", "", "", ssn,  # Comp Code, Minority, Veteran, Home Acct, GL Acct, GL SubAcct, SSN
        "", "",                    # Uploaded, Upload Message
        "", "", "", "",            # 4 empty columns
    ]
    contractors.append(row_data)

# ── Build workbook replicating all 3 sheets ───────────────────────────────────

wb_new = Workbook()

# ── Sheet 1: Instructions ──────────────────────────────────────────────────────
ws_instr = wb_new.active
ws_instr.title = "Instructions"

instructions_data = [
    ("COLUMN NAME",       "MANDATORY",                    "FORMAT"),
    ("First Name",        "TRUE",                         None),
    ("Middle Name",       "",                             None),
    ("Last Name",         "TRUE",                         None),
    ("Job Title",         "TRUE",                         None),
    ("Employee Code",     "",                             None),
    ("Gender",            "",                             "Must be one of these:- MALE, FEMALE, NON-BINARY, PREFER-NOT-TO-DISCLOSE"),
    ("Birth Date",        "TRUE",                         "YYYY-MM-DD , MM-YYYY-DD, or MM-DD-YYYY"),
    ("Hire Date",         "TRUE",                         "YYYY-MM-DD , MM-YYYY-DD, or MM-DD-YYYY"),
    ("Start Date",        "TRUE",                         "YYYY-MM-DD , MM-YYYY-DD, or MM-DD-YYYY"),
    ("Termination Date",  "",                             "YYYY-MM-DD , MM-YYYY-DD, or MM-DD-YYYY"),
    ("Role",              "TRUE",                         "Must be one of these:- Admin, Foreman, Worker; Any other managerial roles must be uploaded as Foreman and can be changed later from the Users profile."),
    ("Cell Phone Number", "TRUE",                         "Must be in the format +(Country Code)(Phone Number), eg +15551234567"),
    ("Email",             "Foreman and Admin Only",       None),
    ("Home Address",      "",                             None),
    ("Employment Type",   "TRUE",                         None),
    ("Active",            "TRUE; Default value: True",    None),
    ("Send Sms",          "TRUE; Default value: False",   None),
    ("Payroll Enabled",   "TRUE; Default value: False",   None),
    ("Compensation Type", "Payroll Enabled Only",         None),
    ("Pay Frequency",     "Payroll Enabled Only",         None),
    ("Compensation Rate", "Payroll Enabled Only",         None),
    ("Overtime Rate",     "Payroll Enabled Only",         None),
    ("Lumber Id",         "",                             None),
    ("External Id",       "",                             None),
    ("Branch Code",       "",                             None),
    ("Branch Name",       "",                             None),
    ("Union Name",        "",                             None),
    ("Job Level",         "",                             None),
    ("Job Classification","",                             None),
    ("Comp Code",         "",                             "classCode-state (e.g. 1234-CA). Multiple values comma-separated."),
    ("Minority",          "",                             None),
    ("Veteran Status",    "",                             None),
    ("Home Account",      "",                             None),
    ("GL Account",        "",                             None),
    ("GL SubAccount",     "",                             None),
    ("SSN",               "FALSE",                        "Must contain exactly 9 digits (no dashes or spaces). Masked values will not be modified. Cannot be cleared once assigned."),
    ("Uploaded",          "",                             None),
    ("Upload Message",    "",                             None),
]
for row in instructions_data:
    ws_instr.append(list(row))

# ── Sheet 2: Employee Info ─────────────────────────────────────────────────────
ws_emp = wb_new.create_sheet("Employee Info")

header = [
    "First Name","Middle Name","Last Name","Job Title","Employee Code","Gender",
    "Birth Date","Hire Date","Start Date","Termination Date","Role","Cell Phone Number",
    "Email","Home Address","Employment Type","Active","Send Sms","Payroll Enabled",
    "Compensation Type","Pay Frequency","Compensation Rate","Overtime Rate",
    "Lumber Id","External Id","Branch Code","Branch Name","Union Name","Job Level",
    "Job Classification","Comp Code","Minority","Veteran Status","Home Account",
    "GL Account","GL SubAccount","SSN","Uploaded","Upload Message",
    None, None, None, None,
]
ws_emp.append(header)

for row_data in contractors:
    ws_emp.append(row_data)


# ── Save ───────────────────────────────────────────────────────────────────────
import os
base = f"contractors_{total}"
out_file = f"{base}.xlsx"
counter = 1
while os.path.exists(out_file):
    try:
        with open(out_file, 'a'):
            pass
        break
    except PermissionError:
        out_file = f"{base}_{counter}.xlsx"
        counter += 1

wb_new.save(out_file)

freq_count = Counter(c[19] for c in contractors)
print()
print(f"Saved: {out_file}")
print(f"Generated {len(contractors)} contractors (rows 2 to {len(contractors)+1})")
print(f"Pay Frequency  : {dict(freq_count)}")
