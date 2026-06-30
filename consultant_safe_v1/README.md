# Consultant-safe EBX-Q Template v1

## Mục đích
Bộ này dành cho consultant nhập liệu ESG định tính theo EBX-Q. Đây là bản an toàn được dựng từ source `2026-06-15`, đã loại cột J chứa dữ liệu thực tế của công ty cùng ngành và đã bổ sung hướng dẫn riêng theo 4 quy mô.

## Cách dùng nhanh
1. Chọn workbook theo ngành và quy mô doanh nghiệp.
2. Vào sheet `EBX-Q 템플릿`, dùng `빈칸초안` làm câu khung và điền mọi vị trí `[ ]` bằng dữ liệu nội bộ.
3. Đọc `작성지침` và `규모지침` trước khi nhập; `규모지침` khác nhau cho `대기업`, `중견`, `중소`, `비상장`.
4. Với chỉ số chưa thống kê, ghi `미집계` hoặc `해당사항 없음`; không để trống, không ước lượng, không điền `0` thay cho dữ liệu thiếu.
5. Chỉ dùng `동종산업 예시` và `동종산업 재서술 예시` để tham khảo cấu trúc diễn đạt, không copy như dữ liệu của doanh nghiệp.
6. Trước khi gửi reviewer, chạy checklist trong `CHECKLIST.md`.

## Nội dung phát hành
- 44 workbook consultant-safe, đủ 11 sector x 4 quy mô.
- Cột thực dữ liệu `동종 실제사례(완성문·회사명 마스킹·수치유지)` đã bị loại.
- Cột `규모지침` được sinh riêng theo từng quy mô và được QA chống trùng giữa 4 quy mô trong cùng sector.
- Các workbook sector SV có cảnh báo mapping confidence thấp/fallback.
- Các mục từng bị review borderline có cảnh báo để consultant dùng ví dụ thận trọng.

## Thống kê build
- Workbooks: 44
- Source JSON: 44
- Scale guidance variants checked: 11
- QA fatal issues: 0
- QA warnings: 0
