# devops-1822040447

BÀI TẬP CÁ NHÂN: TRIỂN KHAI WEBSITE TỰ ĐỘNG VỚI GITHUB ACTIONS

## 1. Thông tin sinh viên
- **Họ và tên**: Phạm Bảo Phong
- **MSSV**: 1822040447
- **Môn học**: Thiết Kế Phần Mềm (TKPM)
- **Giảng viên hướng dẫn**: ThS. Bùi Minh Phụng
- **Địa điểm**: Biên Hòa, Đồng Nai

## 2. Mục tiêu dự án
- Thực hành quy trình CI/CD cơ bản: Code -> Push GitHub -> Kiểm tra tự động -> Tự động Deploy.
- Áp dụng GitHub Actions pipeline để kiểm soát chất lượng (Job `check`) trước khi deploy lên GitHub Pages (Job `deploy`).

## 3. Cấu trúc thư mục
```
├── index.html                # Trang web chính (chứa họ tên, MSSV, title)
├── style.css                 # Định dạng giao diện responsive
├── script.js                 # Xử lý sự kiện và đồng hồ hệ thống
├── .github/
│   └── workflows/
│       └── deploy.yml        # CI/CD Pipeline (job check & job deploy)
└── README.md
```

## 4. Đường dẫn triển khai công khai
- **Website Live**: [https://phong-pham-bao.github.io/devops-1822040447/](https://phong-pham-bao.github.io/devops-1822040447/)
