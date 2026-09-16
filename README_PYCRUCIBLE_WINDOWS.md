# HR Attendance Dashboard + PyCrucible (Windows)

This package is prepared to build the Streamlit HR Attendance Dashboard into a single Windows executable using PyCrucible.

## What HR users receive

`HR_Attendance_Dashboard.exe`

The HR employee double-clicks the EXE. It starts a local Streamlit server on `127.0.0.1`, automatically opens the dashboard in the default browser, and keeps the local app running until the executable is closed.

No Python installation is required on the HR PC.

## 1. Build PC requirements

Use a Windows build machine for the Windows EXE. Install:

```powershell
py -m pip install --upgrade pip
py -m pip install pycrucible
```

Then verify:

```powershell
pycrucible --version
```

PyCrucible's documentation states that the project can be embedded with:

```powershell
pycrucible -e . -o .\dist\HR_Attendance_Dashboard.exe
```

## 2. Build

From this folder:

```powershell
.\build_windows.ps1
```

or:

```bat
build_windows.bat
```

The output is:

```text
dist\HR_Attendance_Dashboard.exe
```

## 3. End-user installation

Copy only:

```text
HR_Attendance_Dashboard.exe
```

to the HR employee's PC.

They can place it on Desktop or a company application folder and double-click it.

## 4. Internet requirement

The normal PyCrucible source-project mode uses `uv` to resolve/install Python dependencies at first launch. Therefore, the first launch should have access to the Python package index or to your approved internal Python package mirror.

After dependencies are prepared, subsequent runs are much faster.

For a completely offline corporate deployment, use an internal package mirror/cache strategy or an embedded wheel-based distribution after testing that workflow in your environment.

## 5. Security / corporate deployment

The dashboard binds only to:

```text
127.0.0.1
```

so it is intended to be accessed only from the local HR PC, not exposed to the LAN.

For production distribution:

- Build from a clean Windows build machine.
- Pin dependencies in `pyproject.toml`.
- Test the EXE in a clean Windows VM.
- Code-sign the final EXE with the company's certificate.
- Distribute the signed EXE through the company's approved software/file distribution mechanism.

## 6. Recommended operating model

The simplest rollout is:

```text
HR Shared Folder / SharePoint / File Server
            |
            +-- HR_Attendance_Dashboard.exe
            |
            +-- Master Karyawan.xlsx
            |
            +-- Template Cuti Karyawan.xlsx
```

Each HR employee runs the EXE locally, while raw attendance exports, leave files, and master employee data can be stored in the company's approved shared location.

## 7. Updating the application

For controlled corporate deployment, release a new versioned EXE, for example:

```text
HR_Attendance_Dashboard_1.0.0.exe
HR_Attendance_Dashboard_1.1.0.exe
HR_Attendance_Dashboard_1.2.0.exe
```

Do not rely on a public GitHub repository for confidential HR source code. PyCrucible documents its GitHub auto-update capability for public repositories; for HR data/software, an internal release/distribution process is preferable.
