---
name: configure-et6-mmo-startup
description: Configure ET6 C# MMO server startup Excel files from server lists, machine/IP data, zone/database settings, process deployment plans, scene topology, merge rules, and load-balance strategy, including reusable placement conventions for routing, global, zone, and load-balance-node servers. Use when Codex needs to create, clone, or update Start*Config@s.xlsx startup configuration workbooks for an ET6 framework MMO server.
---

# Configure ET6 MMO Startup

## Quick Start

Use this skill to produce a complete startup configuration set from a user-provided server list and deployment plan. Work from an existing template directory containing the six startup workbooks:

- `StartMachineConfig@s.xlsx`
- `StartProcessConfig@s.xlsx`
- `StartSceneConfig@s.xlsx`
- `StartZoneConfig@s.xlsx`
- `StartBalanceConfig@s.xlsx`
- `StartMergeServerConfig@s.xlsx`

If the user does not provide a template directory, first check the current working directory for the six expected startup workbooks. Do not hardcode a previous template or output directory name; deployment folders vary between tasks.

Before editing any workbook, read `references/startup-config-schema.md`.

## Reusable Script

Use `scripts/generate_startup_config.py` when the request fits the standard role-based deployment model. Provide it a JSON spec with machine roles, zones, node counts, and database settings. Inspect and patch the generated output when the user provides custom service placement, nonstandard ID ranges, or merge/load-balance behavior not represented by the script. The script is a reusable accelerator, not a substitute for workbook inspection and validation.

## Workflow

1. Collect inputs: target output directory/name, machine list, inner/outer IPs, zone IDs, database connection strings, process-to-machine placement, port ranges, game-zone display names, backend zone IDs, and any merge/load-balance rules.
2. Inspect all six template workbooks before writing. Preserve workbook names, sheet names, row/column positions, formulas, formats, hidden rows, and data validation.
3. Treat rows 3-5 as the schema block in each sheet: row 3 is Chinese label text, row 4 is the English field key, and row 5 is the ET type. Data starts at row 6 unless inspection proves otherwise.
4. Use the English field keys for mapping. Do not rely on localized label text because console encoding may display it incorrectly.
5. Generate or update data rows only. Keep the first two blank leading columns and the schema rows intact.
6. Cross-check IDs and ports across workbooks before delivering the files.

## Workbook Rules

- `StartMachineConfig@s.xlsx`: define one row per physical or logical deployment machine. `Id` is referenced by process placement.
- `StartProcessConfig@s.xlsx`: define process IDs, `MachineId`, and `InnerPort`. Each scene `Process` must exist here.
- `StartSceneConfig@s.xlsx`: define global scenes, routing/node services, and game-zone scenes. Each `Id` must be unique across all sheets.
- `StartZoneConfig@s.xlsx`: define MongoDB database zones. Game-zone scene rows reference zone and backend IDs.
- `StartBalanceConfig@s.xlsx`: define load-balance strategy by `SceneType`.
- `StartMergeServerConfig@s.xlsx`: define merge relationships only when the deployment needs merge-server behavior.

Prefer copying and adapting existing rows with the same `SceneType` rather than inventing new values. Preserve numeric identifiers as numbers unless the template stores that column as text for a reason.

## Standard Deployment Conventions

When the user provides a high-level server list like routing machines, zone machines, load-balance machines, and global machines, apply these defaults unless the user overrides them:

- `StartMachineConfig@s.xlsx`: set `InnerIP` and `OuterIP` to the machine's internal IP. Set `OutRealIP` to the public IP when present, otherwise to the internal IP. Set `EnableGameZone=1` only for physical machines that deploy game-zone services; set it to `0` for routing, load-balance, and global-only machines.
- Routing machines: deploy `RoutingNodeServer` processes on routing machines, with externally exposed `OuterPort` values. Deploy the singleton `RouterManager` on the first routing machine unless the user names a different host.
- Global machines: deploy all global services on the global machine. Treat `RealmBalanceServer` (gateway and logic load balance manager), `BattleBalanceServer`, and `MapNodeBalanceServer` as global-machine services, not load-balance-node-machine services.
- Load-balance machines: deploy node services such as `LogicNodeServer`, `GateNodeServer`, `BattleNodeServer`, and `MapNodeServer` on the load-balance machine.
- Zone machines: deploy all per-zone game services for requested zones on the zone machine and keep `LinkMachineId` aligned with the zone machine.
- Machine control: create one `MachineControlServer` scene and process for every physical server, with `LinkMachineId` equal to that physical machine ID.

## Generation Rules

- Maintain deterministic ID allocation. Respect the template ranges documented in `references/startup-config-schema.md`.
- Keep `Process` equal to a valid process ID from `StartProcessConfig@s.xlsx`.
- Keep `MachineId` equal to a valid machine ID from `StartMachineConfig@s.xlsx`.
- Avoid duplicate `Id`, `Name`, `InnerPort`, or externally exposed `OuterPort` values unless the existing template intentionally duplicates a value.
- Generate `Name` values consistently with the template, such as `GateNodeServer11500` for node scenes and `StatusServer-1` for zone scenes.
- For game-zone rows, keep `InnerZoneIndex`, `BackendZoneId`, `GameZoneDisplayName`, and open/close timestamps aligned across all services for the same logical zone.
- Leave optional ports blank or `0` according to the template's pattern for that `SceneType`.

## Validation

Before finalizing:

- Confirm all six expected workbooks exist in the output.
- Confirm every `StartSceneConfig@s.xlsx` `Process` exists in `StartProcessConfig@s.xlsx`.
- Confirm every `StartProcessConfig@s.xlsx` `MachineId` exists in `StartMachineConfig@s.xlsx`.
- Confirm IDs are unique in each workbook and scene IDs are unique across all scene sheets.
- Confirm required ports do not collide on the same machine.
- Confirm `RealmBalanceServer`, `BattleBalanceServer`, and `MapNodeBalanceServer` processes are assigned to the global machine when using the standard placement conventions.
- Confirm only zone-service machines have `EnableGameZone=1`, and `OuterIP` matches `InnerIP` in `StartMachineConfig@s.xlsx`.
- Confirm database connection values and credentials are filled for every required zone.
- Open or inspect the resulting `.xlsx` files to verify schema rows and formatting were preserved.

When reporting completion, list the output directory and summarize the generated machines, processes, scenes, zones, and any validation warnings.
