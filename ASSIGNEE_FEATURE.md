# Assignee Distribution Feature

## Overview
This feature allows you to pre-assign experiences to different team members when importing ticket sheets. It includes:
- Data validation dropdown for the Assignee column
- Interactive UI for distributing workload among assignees
- Two distribution modes: Manual and Even split

## How It Works

### Web Application Flow:
1. **Upload CSV**: Upload your ticket CSV file
2. **Assignee Distribution Page**: Configure how to assign experiences
   - Select which assignees should receive work
   - Choose between manual counts or even distribution
   - See real-time summary of assignments
3. **Process & Download**: Get your organized XLSX with assignees populated

### Distribution Modes:

#### Manual Assignment
- Set custom experience counts for each assignee
- Full control over workload distribution
- Can assign different amounts to different people

#### Distribute Evenly
- Automatically splits experiences equally among selected assignees
- Handles remainders fairly (first N assignees get +1 experience)
- Just check assignees and click "Distribute Evenly"

## Available Assignees
- Cs. Zsófi
- Ricsi
- Attila
- Lili
- Flóra
- Csabi
- Marci
- Gábor
- Bálint
- B.Zsófi
- Muki
- Péter
- Adri
- Era
- Endre
- Vencel
- Áron

## XLSX Features
- **Assignee Column**: Dropdown data validation with all available assignees
- **Pre-filled Values**: Experiences are automatically assigned based on your distribution settings
- **Editable**: You can still manually change assignees in Excel after download

## Technical Details

### Assignment Logic:
1. Experiences are sorted by Experience ID (low to high)
2. Only experiences with valid Experience IDs are assigned
3. Assignments are distributed round-robin style based on configured counts
4. All tickets in the same experience group get the same assignee
5. Experiences without IDs remain unassigned

### Example:
If you set:
- Ricsi: 3 experiences
- Attila: 2 experiences

The first 3 experiences go to Ricsi, next 2 to Attila, then cycles back (exp 6 → Ricsi, etc.)

## URLs
- Upload page: http://localhost:5001/
- Assignee distribution: http://localhost:5001/assignee-distribution (after upload)

## Files Modified
- `app.py`: Added assignee distribution routes and assignment logic
- `templates/assignee_distribution.html`: New UI page for workload distribution
- `ticket_organizer.py`: Updated to use consistent ASSIGNEES list
