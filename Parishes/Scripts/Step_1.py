import pandas as pd
import re

# ============================================================
# SETTINGS
# ============================================================

INPUT_FILE = "Parishes/Input/Delagates_Names.xlsx" 
OUTPUT_FILE = "Parishes/Processed_Data_Cleaned.xlsx"


# ============================================================
# OFFICIAL ARCHDIOCESE VICARIATE / PARISH LIST
# ============================================================

archdioceseData = [
    {
        "name": "Vicariate of St. Vincent Ferrer",
        "parishes": [
            "Santo Domingo Ybañez de Erquicia Parish",
            "Archdiocesan Shrine of St. Vincent Ferrer",
            "Saint Catherine of Siena Parish",
            "Saint John the Baptist Parish",
            "San Lorenzo Ruiz Parish",
            "Mary Help of Christians Parish",
            "San Isidro Labrador Parish"
        ]
    },
    {
        "name": "Vicariate of Sts. Peter and Paul",
        "parishes": [
            "Saints Peter and Paul Parish",
            "Cristo Divino Tesoro Parish",
            "Our Lady of Fatima Parish",
            "St. Joseph of Nazareth Parish",
            "Santuario de Santa Barbara Holy Family Parish",
            "St. Joseph the Worker Parish",
            "San Isidro a Dumaralos Parish"
        ]
    },
    {
        "name": "Vicariate of St. John the Evangelist",
        "parishes": [
            "St. Louis-Marie Grignion de Montfort Parish",
            "Most Holy Rosary Parish",
            "St. John the Evangelist Cathedral Parish",
            "Annunciation of the Lord Parish",
            "St. Gabriel the Archangel Parish"
        ]
    },
    {
        "name": "Vicariate of Our Lady of the Rosary",
        "parishes": [
            "St. Joseph the Patriarch Parish",
            "Holy Cross Parish",
            "Our Lady of the Blessed Sacrament Parish",
            "Our Lady of the Rosary Parish"
        ]
    },
    {
        "name": "Vicariate of Saint Dominic de Guzman",
        "parishes": [
            "Saint Francis Gil de Federich Chaplaincy",
            "Our Lady of Consolation Parish",
            "Saint John Paul II Parish",
            "Our Lady of Lourdes Pastoral Station",
            "Saint Pius V Parish",
            "San Padre Pio Parish",
            "Most Holy Trinity Parish",
            "Minor Basilica of Saint Dominic de Guzman Parish",
            "Holy Family Parish",
            "St. Joseph, Husband of Mary Parish",
            "San Pedro Calungsod Parish"
        ]
    },
    {
        "name": "Vicariate of St. Ildephonse",
        "parishes": [
            "Our Mother of Perpetual Help Parish",
            "San Antonio de Padua Parish",
            "Jesus the Nazarene Pastoral Station",
            "Saint Francis of Assisi Parish",
            "St. Ildephonse Parish"
        ]
    },
    {
        "name": "Vicariate of Epiphany of Our Lord",
        "parishes": [
            "Epiphany of Our Lord Co-Cathedral Parish",
            "Immaculate Heart of Mary Parish",
            "Resurrection of the Lord",
            "St. Rose of Lima Parish"
        ]
    },
    {
        "name": "Vicariate of Our Lady of Purification",
        "parishes": [
            "Our Lady of Purification Parish",
            "Jesus the Nazarene Parish",
            "Divine Mercy Parish",
            "San Roque Parish",
            "San Isidro Labrador Pastoral Station"
        ]
    },
    {
        "name": "Vicariate of St. Thomas Aquinas",
        "parishes": [
            "Saint Thomas Aquinas Parish",
            "St. Fabian, Pope & Martyr Parish",
            "Beato Juan Martinez de Sto. Domingo Pastoral Station",
            "Virgen de la Medalla Milagrosa Parish",
            "Our Lady of the Miraculous Medal Parish",
            "Saint Hyacinth Parish",
            "St. Jude Thaddeus Parish",
            "Divine Mercy Pastoral Station"
        ]
    }
]


# ============================================================
# READ EXCEL FILE
# ============================================================

df = pd.read_excel(INPUT_FILE)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_value(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def clean_text(value):

    value = clean_value(value)

    if not value:
        return ""

    value = re.sub(r"\s+", " ", value)

    value = value.lower().title()

    replacements = {
        "St.": "St.",
        "Br.": "Br.",
        "Br": "Br.",
        "Fr.": "Fr.",
        "Fr": "Fr.",
        "Sr.": "Sr.",
        "Sr": "Sr.",
        "Jr.": "Jr.",
        "Jr": "Jr.",
        "Ii": "II",
        "Iii": "III",
        "Iv": "IV",
        "V": "V",
        "Ph": "PH",
        "Pym": "PYM",
        "Cym": "CYM",
        "Soccom": "SocCom",
    }

    for old, new in replacements.items():

        value = re.sub(
            rf"\b{re.escape(old)}\b",
            new,
            value
        )

    return value.strip()


def clean_parish(value):

    value = clean_value(value)

    if not value:
        return ""

    parishes = value.split(";")

    cleaned_parishes = []

    for parish in parishes:

        parish = clean_text(parish)

        if parish:
            cleaned_parishes.append(parish)

    return "; ".join(cleaned_parishes)


# ============================================================
# FIND PARISH COLUMNS
# ============================================================

parish_columns = [
    col for col in df.columns
    if col.startswith(
        "Select your Parish / Church Affiliation"
    )
]


# ============================================================
# FIND PARTICIPANTS
# ============================================================

participant_numbers = []

for col in df.columns:

    match = re.match(
        r"Participant (\d+) Full Name",
        col
    )

    if match:

        participant_numbers.append(
            int(match.group(1))
        )


participant_numbers = sorted(
    participant_numbers
)


print(
    "Participants detected:",
    participant_numbers
)

print(
    "Parish columns detected:",
    len(parish_columns)
)


# ============================================================
# CREATE CLEANED RECORDS
# ============================================================

records = []


for _, row in df.iterrows():

    timestamp = clean_value(
        row.get("Timestamp")
    )

    email = clean_value(
        row.get("Email Address")
    )

    vicariate = clean_text(
        row.get("Select Your Vicariate")
    )

    delegation_head_name = clean_text(
        row.get("Delegation Head Full Name")
    )

    delegation_head_nickname = clean_text(
        row.get("Delegation Head Nickname")
    )

    delegation_head_age = clean_value(
        row.get("Delegation Head Age")
    )

    delegation_head_size = clean_value(
        row.get("Delegation Head T-Shirt Size")
    )

    contact_number = clean_value(
        row.get("Delegation Head Contact Number")
    )


    # ========================================================
    # GET PARISHES
    # ========================================================

    parishes = []

    for parish_column in parish_columns:

        parish = clean_parish(
            row.get(parish_column)
        )

        if parish and parish not in parishes:

            parishes.append(parish)


    if not parishes:

        parishes = [""]


    parish_text = "; ".join(parishes)


    # ========================================================
    # DELEGATION HEAD
    # ========================================================

    if delegation_head_name:

        records.append({

            "Type": "Delegation Head",

            "Timestamp": timestamp,

            "Email": email,

            "Vicariate": vicariate,

            "Parish / Church": parish_text,

            "Full Name": delegation_head_name,

            "Nickname": delegation_head_nickname,

            "Age": delegation_head_age,

            "T-Shirt Size": delegation_head_size,

            "Contact Number": contact_number,

        })


    # ========================================================
    # PARTICIPANTS
    # ========================================================

    for participant_number in participant_numbers:

        name_column = (
            f"Participant {participant_number} Full Name"
        )

        nickname_column = (
            f"Participant {participant_number} Nickname"
        )

        age_column = (
            f"Participant {participant_number} Age"
        )

        size_column = (
            f"Participant {participant_number} T-Shirt Size"
        )


        name = clean_text(
            row.get(name_column)
        )


        if not name:
            continue


        nickname = clean_text(
            row.get(nickname_column)
        )

        age = clean_value(
            row.get(age_column)
        )

        shirt_size = clean_value(
            row.get(size_column)
        )


        records.append({

            "Type": "Participant",

            "Timestamp": timestamp,

            "Email": "",

            "Vicariate": vicariate,

            "Parish / Church": parish_text,

            "Full Name": name,

            "Nickname": nickname,

            "Age": age,

            "T-Shirt Size": shirt_size,

            "Contact Number": "",

        })


# ============================================================
# CREATE CLEANED DATAFRAME
# ============================================================

cleaned_df = pd.DataFrame(records)


cleaned_df.insert(
    0,
    "Person No.",
    range(
        1,
        len(cleaned_df) + 1
    )
)


# ============================================================
# ============================================================
# CHECK SUBMISSION STATUS
# ============================================================
# ============================================================

status_records = []


# ------------------------------------------------------------
# CREATE NORMALIZED LOOKUP FUNCTIONS
# ------------------------------------------------------------

def normalize_for_comparison(value):

    value = clean_value(value)

    value = value.lower()

    # Normalize common spelling/formatting differences
    value = value.replace("st.", "st")
    value = value.replace("sts.", "sts")
    value = value.replace("&", "and")

    # Remove punctuation
    value = re.sub(
        r"[^a-z0-9\s]",
        "",
        value
    )

    # Normalize spaces
    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


# ------------------------------------------------------------
# BUILD SUBMITTED VICARIATE / PARISH SET
# ------------------------------------------------------------

submitted_vicariates = set()
submitted_parishes = set()


for _, row in cleaned_df.iterrows():

    vicariate = normalize_for_comparison(
        row["Vicariate"]
    )

    if vicariate:

        submitted_vicariates.add(
            vicariate
        )


    parish_data = clean_value(
        row["Parish / Church"]
    )

    if parish_data:

        for parish in parish_data.split(";"):

            parish = normalize_for_comparison(
                parish
            )

            if parish:

                submitted_parishes.add(
                    parish
                )


# ============================================================
# CHECK EACH OFFICIAL VICARIATE
# ============================================================

for vicariate_data in archdioceseData:

    official_vicariate = vicariate_data["name"]

    official_vicariate_key = (
        normalize_for_comparison(
            official_vicariate
        )
    )


    # --------------------------------------------------------
    # Check whether vicariate submitted anything
    # --------------------------------------------------------

    vicariate_submitted = (
        official_vicariate_key
        in submitted_vicariates
    )


    # --------------------------------------------------------
    # Check individual parishes
    # --------------------------------------------------------

    for parish in vicariate_data["parishes"]:

        parish_key = normalize_for_comparison(
            parish
        )

        parish_submitted = (
            parish_key
            in submitted_parishes
        )


        if parish_submitted:

            status = "SUBMITTED"

        else:

            status = "NO PARTICIPANTS"


        status_records.append({

            "Vicariate": official_vicariate,

            "Parish": parish,

            "Vicariate Status":
                "SUBMITTED"
                if vicariate_submitted
                else "NO SUBMISSION",

            "Parish Status": status,

        })


# ============================================================
# CREATE STATUS DATAFRAME
# ============================================================

status_df = pd.DataFrame(
    status_records
)


# ============================================================
# ADD VICARIATE SUMMARY
# ============================================================

vicariate_summary = []


for vicariate_data in archdioceseData:

    vicariate = vicariate_data["name"]

    vicariate_rows = status_df[
        status_df["Vicariate"]
        == vicariate
    ]


    total_parishes = len(
        vicariate_rows
    )

    submitted = len(
        vicariate_rows[
            vicariate_rows["Parish Status"]
            == "SUBMITTED"
        ]
    )

    missing = total_parishes - submitted


    vicariate_summary.append({

        "Vicariate": vicariate,

        "Total Parishes":
            total_parishes,

        "Submitted Parishes":
            submitted,

        "Missing Parishes":
            missing,

        "Status":
            "COMPLETE"
            if missing == 0
            else "INCOMPLETE"

    })


vicariate_summary_df = pd.DataFrame(
    vicariate_summary
)


# ============================================================
# SAVE EVERYTHING INTO ONE EXCEL FILE
# ============================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    # Cleaned participants
    cleaned_df.to_excel(
        writer,
        index=False,
        sheet_name="Participants"
    )


    # Detailed parish status
    status_df.to_excel(
        writer,
        index=False,
        sheet_name="Submission Status"
    )


    # Vicariate summary
    vicariate_summary_df.to_excel(
        writer,
        index=False,
        sheet_name="Vicariate Summary"
    )


# ============================================================
# CONSOLE SUMMARY
# ============================================================

total_vicariates = len(
    archdioceseData
)

total_parishes = sum(
    len(v["parishes"])
    for v in archdioceseData
)

submitted_parishes = len(
    status_df[
        status_df["Parish Status"]
        == "SUBMITTED"
    ]
)

missing_parishes = (
    total_parishes
    - submitted_parishes
)

submitted_vicariates = len(
    vicariate_summary_df[
        vicariate_summary_df["Status"]
        != "INCOMPLETE"
    ]
)


print()
print(
    "=============================================="
)

print(
    "CLEANING + SUBMISSION CHECK COMPLETE"
)

print(
    "=============================================="
)

print()

print(
    f"Total people extracted: "
    f"{len(cleaned_df)}"
)

print(
    "Delegation Heads:",
    (
        cleaned_df["Type"]
        == "Delegation Head"
    ).sum()
)

print(
    "Participants:",
    (
        cleaned_df["Type"]
        == "Participant"
    ).sum()
)

print()

print(
    f"Total Vicariates: "
    f"{total_vicariates}"
)

print(
    f"Total Parishes: "
    f"{total_parishes}"
)

print(
    f"Parishes with participants: "
    f"{submitted_parishes}"
)

print(
    f"Parishes without participants: "
    f"{missing_parishes}"
)

print()

print(
    "Output file:",
    OUTPUT_FILE
)

print()


# ============================================================
# PRINT MISSING PARISHES
# ============================================================

print(
    "=============================================="
)

print(
    "PARISHES WITHOUT PARTICIPANTS"
)

print(
    "=============================================="
)

for vicariate_data in archdioceseData:

    vicariate = vicariate_data["name"]

    missing = status_df[
        (status_df["Vicariate"] == vicariate)
        &
        (
            status_df["Parish Status"]
            == "NO PARTICIPANTS"
        )
    ]


    if not missing.empty:

        print()
        print(vicariate)

        for parish in missing["Parish"]:

            print(
                f"  - {parish}"
            )

print()

print(
    "=============================================="
)

print(
    "END"
)

print(
    "=============================================="
)