# TK Onluyen - Transfer Dữ Liệu

## Cài đặt & Chạy

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cấu trúc thư mục

```
onluyen_app/
├── app.py              # Giao diện Streamlit
├── core.py             # Logic xử lý & gán dữ liệu
├── requirements.txt
└── templates/
    ├── TK Onluyen (Admin, GV, HS).xlsx
    └── TK Onluyen (Tiểu học).xlsx
```

## Quy tắc file data input

| Sheet trong file data | Dữ liệu lấy | Gán vào sheet template |
|---|---|---|
| `GIAO-VIEN` | A2:E (bỏ dòng trống) | GIAO-VIEN → A2:E |
| `Danh sách HS toàn trường` | A7:I (bỏ dòng trống) | HOC-SINH → A2:I |

## Tên file export

| Loại | Tên file |
|---|---|
| Gộp chung | `TK Onluyen (Admin, GV, HS) - Tên trường.xlsx` |
| Tách lẻ | `TK Onluyen - ADMIN/GIAO-VIEN/HOC-SINH - Tên trường.xlsx` (zip) |
| Tiểu học | `TK Onluyen (Tiểu học) - Tên trường.xlsx` |
