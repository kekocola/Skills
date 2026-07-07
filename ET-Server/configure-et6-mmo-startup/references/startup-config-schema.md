# ET6 MMO Startup Config Schema

Use this reference after loading `$configure-et6-mmo-startup`.

## Template Layout

All startup files are Excel workbooks. The observed template layout is:

- Columns A-B are blank leading columns.
- Row 3 contains human labels.
- Row 4 contains English field keys.
- Row 5 contains ET field types.
- Row 6 and below contain data rows.

Use row 4 keys as the source of truth for column mapping. Preserve row 3 labels and row 5 types.

## Workbook Summary

| Workbook | Sheet pattern | Primary fields |
| --- | --- | --- |
| `StartMachineConfig@s.xlsx` | `StartMachineConfig` | `Id`, `InnerIP`, `OuterIP`, `OutRealIP`, `EnableGameZone` |
| `StartProcessConfig@s.xlsx` | `StartProcessConfig` | `Id`, `MachineId`, `InnerPort` |
| `StartSceneConfig@s.xlsx` | Multiple service-category sheets | `Id`, `Process`, `Zone`, `SceneType`, `SceneDesc`, `Name`, `OuterPort` plus category-specific fields |
| `StartZoneConfig@s.xlsx` | `StartZoneConfig` | `Id`, `DBConnection`, `DBName`, `Account`, `Password`, `Desc` |
| `StartBalanceConfig@s.xlsx` | Load-balance strategy sheet | `Id`, `SceneType`, `Strategy` |
| `StartMergeServerConfig@s.xlsx` | Merge strategy sheet | `Id`, `MainServer`, `SubServer`, `UpdateAt` |

## Scene Sheets and ID Ranges

`StartSceneConfig@s.xlsx` is split by service category. Use the existing sheet and ID range when adding rows:

| Category | ID range | Common use |
| --- | --- | --- |
| Global system services | `10000-10249` | global control, login, router manager |
| Global business services | `10250-10499` | rank, misc, mail, activity-like global services |
| System routing/firewall | `10500-10999` | routing nodes |
| Logic load-balance nodes | `11000-11499` | `LogicNodeServer` |
| Gate load-balance nodes | `11500-11999` | `GateNodeServer` |
| Battle load-balance nodes | `12000-12249` | `BattleNodeServer` |
| Map load-balance nodes | `12300-12399` | `MapNodeServer` |
| System control services | `12500-19999` | machine control and control-plane services |
| Game-zone services | `30000+` | per-zone gameplay services |

## Cross-Workbook Relationships

- `StartSceneConfig@s.xlsx.Process` -> `StartProcessConfig@s.xlsx.Id`.
- `StartProcessConfig@s.xlsx.MachineId` -> `StartMachineConfig@s.xlsx.Id`.
- `StartSceneConfig@s.xlsx.Zone` should align with `StartZoneConfig@s.xlsx.Id` where the scene is zone/database scoped.
- Game-zone service rows should share coherent `InnerZoneIndex`, `BackendZoneId`, open/close timestamps, and display names for the same logical zone.
- `StartBalanceConfig@s.xlsx.SceneType` should reference scene types that need load-balance routing.

## Input Normalization

Normalize user-provided deployment input into these structures before editing:

```text
machines:
  - id
    inner_ip
    outer_ip
    out_real_ip
    enable_game_zone

processes:
  - id
    machine_id
    inner_port

zones:
  - id
    db_connection
    db_name
    account
    password
    desc

scenes:
  - id
    process
    zone
    scene_type
    scene_desc
    name
    outer_port
    optional fields matching the target sheet
```

If the user supplies a higher-level server list, derive rows by following the nearest matching template rows. Ask only for values that cannot be inferred safely, such as database credentials, public IPs, or authoritative zone/backend IDs.

## Editing Guidance

- Use a real spreadsheet parser/writer such as `openpyxl` or the Codex spreadsheet runtime; do not edit `.xlsx` XML by hand.
- Load workbooks with formatting preserved. Copy existing style from neighboring data rows when inserting additional rows.
- Delete stale data rows only when creating a fresh config set and only below the schema rows.
- Preserve workbook encoding and formulas.
- Save generated files to a new output directory unless the user explicitly asks to overwrite the template.

## Validation Queries

Run equivalent checks before final delivery:

- Duplicate IDs in each workbook.
- Missing machine references from processes.
- Missing process references from scenes.
- Port collisions by `(MachineId, InnerPort)` and by externally exposed `OuterPort`.
- Blank required fields in generated rows.
- Scene IDs outside their category range.
- Game-zone rows with inconsistent `BackendZoneId`, `InnerZoneIndex`, or display name for the same zone.
