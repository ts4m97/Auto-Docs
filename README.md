# Auto Docs

Auto Docs la ung dung Windows offline dung Python + PySide6 de tao tai lieu tu mau Word.

## Chuc nang chinh

- Them file mau `.docx` va quet placeholder dang `[[ten_truong]]`.
- Luu mau vao thu vien cuc bo.
- Nhap du lieu thu cong theo cac placeholder da quet.
- Tao file Excel mau tu placeholder.
- Nhap du lieu hang loat tu Excel.
- Xuat file `.docx` theo quy tac dat ten.
- Tuy chon xuat PDF va anh neu may Windows co Microsoft Word va cac thu vien can thiet.
- Mo nhanh thu muc chua file sau khi xuat.
- In hang loat: keo tha file, sap xep thu tu, chon may in, so ban va khoang nghi giua cac file.
- Lich su xuat file: luu thoi gian, dinh dang, file da tao va cac truong du lieu da dien.

## Luu y ve file mau

Khi them mau, Auto Docs copy file `.docx` vao thu vien cuc bo `data/templates`.
Vi vay co the xoa file Word goc sau khi them mau. Mau chi bi mat neu xoa mau trong
ung dung hoac xoa thu muc `data` cua Auto Docs.

## Luu y ve in hang loat

- File Word `.doc`, `.docx`, `.rtf` duoc in qua Microsoft Word de giu bo cuc.
- File PDF, anh, TXT duoc gui qua lenh in mac dinh cua Windows, nen may can co ung
  dung mac dinh ho tro in dinh dang do.
- Ung dung gui file theo thu tu trong danh sach va cho nghi giua tung file de tranh
  day lenh qua nhanh vao may in.

## Lich su xuat file

Moi lan xuat DOCX/PDF/anh, Auto Docs luu lai mot dong lich su trong SQLite cuc bo,
gom mau da dung, dinh dang xuat, duong dan file va toan bo gia tri placeholder tai
thoi diem xuat. Lich su nam trong `data/autodocs.sqlite3`.

## Cai dat

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run.py
```

Du lieu ung dung duoc luu trong thu muc `data/`.

## Chay ban EXE

Ban Windows da build nam tai:

```powershell
dist\AutoDocs\AutoDocs.exe
```

Khi chay ban EXE, du lieu offline se duoc tao canh file `.exe` trong:

```text
dist\AutoDocs\data
dist\AutoDocs\exports
```

## Build lai EXE

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
```

Script nay se tao lai icon, dong metadata Windows va build app khong mo console.
