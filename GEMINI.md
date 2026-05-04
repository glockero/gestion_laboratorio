# GEMINI.md

## Directory Overview
This directory serves as a repository for tracking technical equipment repairs, likely within a laboratory or technical service department. The primary data is stored in a structured CSV format, providing a historical and real-time record of repair activities across different locations.

## Key Files
- **reparaciones.csv**: The central data file containing over 10,000 records of equipment repairs.
    - **Columns**:
        - `fecha`: Date of the repair entry.
        - `sala`: The location or room associated with the equipment (e.g., WILDE, BIYEMAS, REBISCO).
        - `uid` / `npu`: Unique identification or tracking numbers for the equipment.
        - `parte`: The specific part of the equipment being addressed.
        - `equipo`: Description of the equipment (e.g., Power supplies, Mini PCs, Monitors, TVs).
        - `urgente`: Indicates if the repair is a priority (SI/NO).
        - `tecnico`: The name of the technician assigned to the task.
        - `estado`: Current status of the repair (e.g., `PEND. DE REVISION`, `REPARADO`, `EN PRUEBA EN SALA`).
        - `observaciones`: Technical notes or additional details about the repair process.

## Usage
The contents of this directory are intended for:
1. **Data Analysis**: Monitoring repair trends, technician performance, and equipment reliability.
2. **Status Tracking**: Checking the current progress of pending repairs.
3. **Inventory/Service History**: Maintaining a log of all technical interventions performed on specific hardware units.

When interacting with this project, focus on querying or updating the `reparaciones.csv` file to extract insights or reflect new service activities.
