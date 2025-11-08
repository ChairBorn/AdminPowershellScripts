# Generate production-ready scripts and manifests based on the uploaded data
import pandas as pd
import numpy as np
import re
from pathlib import Path
import json
from textwrap import dedent

base = Path("/mnt/data")

# Input files
users_path = base / "BLS-USERS.xlsx"
groups_export_path = base / "exportGroup_2025-10-15.csv"
priority_path = base / "IT Master Priority List - Copy.xlsx"

# Helper loaders
def read_any_excel(path):
    xl = pd.ExcelFile(path)
    sheets = {name: xl.parse(name) for name in xl.sheet_names}
    return sheets

def safe_read_csv(path):
    for encoding in [None, "utf-16", "latin-1", "utf-8-sig"]:
        try:
            return pd.read_csv(path, encoding=encoding) if encoding else pd.read_csv(path)
        except Exception:
            continue
    return pd.read_csv(path, errors="ignore")

def normalize_cols(df):
    df2 = df.copy()
    df2.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df2.columns]
    return df2

# Load users and priority
users_book = read_any_excel(users_path)
user_sheet_name, user_df = max(users_book.items(), key=lambda kv: kv[1].shape[0])
user_df = normalize_cols(user_df)

priority_book = read_any_excel(priority_path)
priority_sheet_name, priority_df = max(priority_book.items(), key=lambda kv: kv[1].shape[0])
priority_df = normalize_cols(priority_df)

# Column mapping for users
col_candidates = { 
    "userPrincipalName": ["userprincipalname","upn","email","signin name","user name","user"],
    "displayName": ["displayname","name","full name"],
    "givenName": ["givenname","first name","firstname"],
    "surname": ["surname","last name","lastname","familyname"],
    "department": ["department","dept"],
    "jobTitle": ["jobtitle","title","position","role"],
    "officeLocation": ["officelocation","location","office","site"],
    "employeeType": ["employeetype","type","worker type","classification"],
    "manager": ["manager","manager upn","manager email","reports to"],
    "employeeId": ["employeeid","id","worker id","personnel number"],
    "companyName": ["company","companyname","tenant"],
    "city": ["city","town"],
    "state": ["state","province","region"],
    "country": ["country","country/region","countryregion"],
    "usageLocation": ["usagelocation","m365 usage location"],
    "mobilePhone": ["mobilephone","mobile","phone","phone number"],
    "licenseSku": ["licensesku","sku","license","assigned license"]
}

def map_columns(df, mapping):
    mapped = {}
    lower_cols = {c.lower(): c for c in df.columns}
    for std, alts in mapping.items():
        found = None
        for a in alts:
            if a in lower_cols:
                found = lower_cols[a]
                break
        mapped[std] = found
    return mapped

mapped_cols = map_columns(user_df, col_candidates)
std_cols = [c for c in mapped_cols.values() if c is not None]
std_df = user_df[std_cols].copy()
rename_map = {mapped_cols[k]: k for k in mapped_cols if mapped_cols[k] is not None}
std_df.rename(columns=rename_map, inplace=True)

# Normalization map
def profile_column(df, col):
    s = df[col].astype(str).fillna("")
    s_norm = s.str.strip()
    variants = {}
    for v in s_norm.unique():
        if not v: 
            continue
        key = v.strip().lower()
        variants.setdefault(key, set()).add(v.strip())
    multi_variant = {k:list(v) for k,v in variants.items() if len(v)>1}
    return multi_variant

attributes_to_normalize = [c for c in ["department","jobTitle","officeLocation","employeeType"] if c in std_df.columns]

norm_rows = []
for attr in attributes_to_normalize:
    multi = profile_column(std_df, attr)
    for key, variants in multi.items():
        canonical = " ".join([w.upper() if w in {"HR","IT","R&D","QA"} else w.capitalize() for w in key.split()])
        for v in variants:
            if v != canonical:
                norm_rows.append({"Attribute": attr, "FromValue": v, "ToValue": canonical})

norm_df = pd.DataFrame(norm_rows).drop_duplicates().sort_values(["Attribute","FromValue"])

# Save normalization CSV
norm_csv_path = base / "Attribute-Normalization-Map.csv"
norm_df.to_csv(norm_csv_path, index=False)

# Build dynamic groups
def make_group_name(prefix, scope, value):
    v = re.sub(r'[^A-Za-z0-9\-\s&/_]', '', str(value)).strip().replace("/", "-")
    v = re.sub(r'\s+', '-', v)
    return f"{prefix}-{scope}-{v}"

dept_values = sorted(set(std_df["department"].dropna().astype(str).str.strip()) - {""}) if "department" in std_df.columns else []
role_keywords = ["Analyst","Manager","Director","Engineer","Specialist","Assistant","Intern","Contractor","Consultant"]

dynamic_groups = []
for d in dept_values[:100]:
    dynamic_groups.append({
        "DisplayName": make_group_name("SG", "Dept", d),
        "Description": f"Department-scoped access for {d}",
        "Rule": f'(user.department -eq "{d}")',
        "SecurityEnabled": True,
        "MailEnabled": False,
        "GroupTypes": ["DynamicMembership"]
    })

for kw in role_keywords:
    dynamic_groups.append({
        "DisplayName": make_group_name("SG", "Role", kw),
        "Description": f"Role-based access: {kw}",
        "Rule": f'(user.jobTitle -contains "{kw}")',
        "SecurityEnabled": True,
        "MailEnabled": False,
        "GroupTypes": ["DynamicMembership"]
    })

# License-based groups if available (informational; creation optional)
license_groups = []
if "licenseSku" in std_df.columns:
    for sku in sorted(set(std_df["licenseSku"].dropna().astype(str).str.strip()) - {""}):
        license_groups.append({
            "DisplayName": make_group_name("LG", "License", sku),
            "Description": f"License assignment group for {sku}",
            "Rule": f'(user.assignedPlans -any (plan.service -eq "{sku}"))',
            "SecurityEnabled": True,
            "MailEnabled": False,
            "GroupTypes": ["DynamicMembership"]
        })

# Access model from priority list
dept_like = next((c for c in priority_df.columns if "department" in c.lower() or "dept" in c.lower()), None)
resource_like = next((c for c in priority_df.columns if any(k in c.lower() for k in ["application","resource","system","app","saas","sharepoint","drive"])), None)
access_like = next((c for c in priority_df.columns if any(k in c.lower() for k in ["access","role","permission","entitlement"])), None)

access_map = []
if dept_like and resource_like:
    cols = [c for c in [dept_like, resource_like, access_like] if c]
    view = priority_df[cols].dropna(how="all").copy()
    view.rename(columns={dept_like:"Department", resource_like:"Resource", (access_like if access_like else "Access"): "Access"}, inplace=True)
    for _, row in view.iterrows():
        d = str(row.get("Department","")).strip()
        r = str(row.get("Resource","")).strip()
        a = str(row.get("Access","")).strip() if "Access" in row else ""
        if not r:
            continue
        # Build an app-role group name
        role_suffix = a if a else "Users"
        grp = make_group_name("SG", f"App-{r}", role_suffix)
        access_map.append({
            "Department": d,
            "Resource": r,
            "AccessRole": role_suffix,
            "AppRoleGroup": grp
        })

access_model = {"generatedFrom": priority_sheet_name, "items": access_map}

# Write app access model
app_model_path = base / "App-Access-Model.json"
with open(app_model_path, "w") as f:
    json.dump(access_model, f, indent=2)

# PowerShell: Create/Update Dynamic Groups via Microsoft Graph
ps_groups_path = base / "Create-Dynamic-Groups.ps1"
ps_groups = dedent(f"""
    <#
    Creates or updates dynamic security groups in Microsoft Entra ID (Azure AD).
    Requires: Microsoft.Graph (Install-Module Microsoft.Graph)
    Sign in:  Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All"
    #>

    param(
        [Parameter(Mandatory=$false)][string]$TenantId
    )

    Import-Module Microsoft.Graph

    if ($TenantId) {{
        Connect-MgGraph -TenantId $TenantId -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All"
    }} else {{
        Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All"
    }}
    Select-MgProfile -Name "beta"

    function Ensure-DynamicGroup {{
        param(
            [string]$DisplayName,
            [string]$Description,
            [string]$Rule,
            [bool]$SecurityEnabled = $true,
            [bool]$MailEnabled = $false,
            [string[]]$GroupTypes = @("DynamicMembership")
        )

        $existing = Get-MgGroup -Filter ("displayName eq '{'{'0'}'}'" -f $DisplayName) -ConsistencyLevel eventual -CountVariable count
        if ($existing) {{
            Write-Host "Updating group: $DisplayName"
            Update-MgGroup -GroupId $existing.Id -Description $Description -MembershipRule $Rule -MembershipRuleProcessingState "On"
        }} else {{
            Write-Host "Creating group: $DisplayName"
            New-MgGroup -DisplayName $DisplayName -Description $Description -MailEnabled:$MailEnabled -SecurityEnabled:$SecurityEnabled -GroupTypes $GroupTypes -MembershipRule $Rule -MembershipRuleProcessingState "On"
        }}
    }}

    # Department groups
""")
# Append group calls
for g in dynamic_groups:
    ps_groups += f"Ensure-DynamicGroup -DisplayName '{g['DisplayName']}' -Description '{g['Description'].replace(\"'\",\"''\")}' -Rule '{g['Rule'].replace(\"'\",\"''\")}'\n"
# (Optional) License groups commented out for safety
if license_groups:
    ps_groups += "\n# License groups (optional — uncomment to enable)\n"
    for g in license_groups:
        ps_groups += f"# Ensure-DynamicGroup -DisplayName '{g['DisplayName']}' -Description '{g['Description'].replace(\"'\",\"''\")}' -Rule '{g['Rule'].replace(\"'\",\"''\")}'\n"

ps_groups += "\nWrite-Host 'Done.'\n"
ps_groups_path.write_text(ps_groups, encoding="utf-8")

# PowerShell: Normalize attributes from CSV map
ps_norm_path = base / "Normalize-Attributes.ps1"
ps_norm = dedent(f"""
    <#
    Normalizes user attributes (Department, JobTitle, OfficeLocation, EmployeeType) based on a CSV mapping.
    CSV expected columns: Attribute,FromValue,ToValue
    Requires: Microsoft.Graph (Install-Module Microsoft.Graph)
    #>

    param(
        [Parameter(Mandatory=$true)][string]$CsvPath,
        [Parameter(Mandatory=$false)][string]$TenantId
    )

    Import-Module Microsoft.Graph

    if ($TenantId) {{
        Connect-MgGraph -TenantId $TenantId -Scopes "User.ReadWrite.All","Directory.AccessAsUser.All"
    }} else {{
        Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.AccessAsUser.All"
    }}
    Select-MgProfile -Name "beta"

    $map = Import-Csv -Path $CsvPath
    $attrs = $map | Group-Object Attribute

    foreach ($group in $attrs) {{
        $attr = $group.Name
        $pairs = $group.Group
        foreach ($pair in $pairs) {{
            $from = $pair.FromValue
            $to = $pair.ToValue
            if ([string]::IsNullOrWhiteSpace($from) -or [string]::IsNullOrWhiteSpace($to)) {{ continue }}

            Write-Host "Normalizing $attr: '$from' -> '$to'"
            # Find users matching the 'from' value (case-insensitive)
            $filterAttr = $attr  # matches Graph property names used earlier
            $users = Get-MgUser -All -Filter "$filterAttr eq '{'{'0'}'}'" -ConsistencyLevel eventual -CountVariable count -ErrorAction SilentlyContinue -Search $null -Property "id,displayName,userPrincipalName,$filterAttr" -Search ""
            if (-not $users) {{ continue }}

            foreach ($u in $users) {{
                try {{
                    $patch = @{{}}
                    $patch[$attr] = $to
                    Update-MgUser -UserId $u.Id -BodyParameter $patch
                    Write-Host "  Updated: $($u.userPrincipalName)"
                }} catch {{
                    Write-Warning "  Failed: $($u.userPrincipalName) - $($_.Exception.Message)"
                }}
            }}
        }}
    }}

    Write-Host "Normalization pass complete."
""")
ps_norm_path.write_text(ps_norm, encoding="utf-8")

# PowerShell: Create App Role Groups from manifest and link feeder groups
ps_app_groups_path = base / "Create-App-Role-Groups.ps1"
ps_app_groups = dedent(r"""
    <#
    Creates app-role security groups per App-Access-Model.json and links feeder groups:
      - SG-Dept-<Dept> and SG-Role-* -> SG-App-<App>-<Role>
    Requires: Microsoft.Graph
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ModelPath,
        [Parameter(Mandatory=$false)][string]$TenantId
    )

    Import-Module Microsoft.Graph
    if ($TenantId) {
        Connect-MgGraph -TenantId $TenantId -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All"
    } else {
        Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.ReadWrite.All"
    }
    Select-MgProfile -Name "beta"

    function Get-OrCreate-Group {
        param([string]$DisplayName,[string]$Description)
        $g = Get-MgGroup -Filter ("displayName eq '{0}'" -f $DisplayName) -ConsistencyLevel eventual -CountVariable count
        if ($g) { return $g }
        return New-MgGroup -DisplayName $DisplayName -Description $Description -MailEnabled:$false -SecurityEnabled:$true
    }

    $model = Get-Content -Raw -Path $ModelPath | ConvertFrom-Json
    $items = $model.items | Sort-Object Resource, AccessRole, Department

    # Create unique app-role groups first
    $appRoles = $items | Group-Object AppRoleGroup
    foreach ($grp in $appRoles) {
        $name = $grp.Name
        $desc = "Application role group"
        $g = Get-OrCreate-Group -DisplayName $name -Description $desc
    }

    # Link feeder groups (dept and role groups) into app-role groups
    foreach ($item in $items) {
        $appGroupName = $item.AppRoleGroup
        $dept = $item.Department
        if
