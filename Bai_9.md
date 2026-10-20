# Bài 9: Tự động tra cứu phòng và thù lao báo cáo viên bằng VLOOKUP

## Phần 1: Phân tích bài toán (I/O)

### 1. Xác định các thành phần của VLOOKUP

Cú pháp:

```
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```


| Thànhphần                  | Nộidung                                                                                                                     |
| ---------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `<span>Lookup_value</span>`  | Mã tác giả tại ô`<span>C2</span>`, dùng `<span>TRIM(C2)</span>`để loại bỏ khoảng trắng thừa ở đầu và cuối. |
| `<span>Table_array</span>`   | Vùng tra cứu`<span>$A$2:$E$81</span>`trên sheet `<span>DanhMucTacGia</span>`, được cố định bằng `<span>$</span>`.  |
| `<span>Col_index_num</span>` | Cột số 2 cho tên tác giả, số 4 cho phòng và số 5 cho thù lao.                                                      |
| `<span>Range_lookup</span>`  | `<span>FALSE</span>`để tìm kiếm chính xác mã tác giả.                                                               |

### 2. Giải pháp xử lý 3 bẫy dữ liệu

* Bẫy 1 – Khoảng trắng:Dùng `<span>TRIM(C2)</span>`để loại bỏ khoảng trắng thừa trong mã tác giả.
* Bẫy 2 – Không tìm thấy mã:Dùng `<span>IFERROR(...,"Chưa </span><span>cập </span><span>nhật")</span>`để thay lỗi `<span>#N/A</span>`bằng thông báo dễ hiểu.
* Bẫy 3 – Vùng tra cứu bị trôi:Dùng `<span>$A$2:$E$81</span>`để cố định vùng tra cứu khi sao chép công thức xuống các dòng khác.

### 3. Các bước thực hiện trong Excel

1. Mở hai sheet `<span>DanhSachBaoCao</span>`và `<span>DanhMucTacGia</span>`trong cùng một file Excel.
2. Xác định mã tác giả nằm ở cột C của sheet `<span>DanhSachBaoCao</span>`.
3. Nhập công thức VLOOKUP kết hợp TRIM và IFERROR vào ô D2.
4. Thay chỉ số cột để tra cứu phòng và thù lao tại E2 và F2.
5. Kéo các công thức xuống hết dòng 81.
6. Kiểm tra các mã tác giả không tồn tại để bảo đảm hiển thị `<span>Chưa </span><span>cập </span><span>nhật</span>`.

## Phần 2: Triển khai và chống lỗi

### 4. Công thức ô D2 – Tên tác giả

```
=IFERROR(VLOOKUP(TRIM(C2),DanhMucTacGia!$A$2:$E$81,2,FALSE),"Chưa cập nhật")
```

Giải thích:

* `<span>TRIM(C2)</span>`: loại bỏ khoảng trắng thừa ở hai đầu mã tác giả.
* `<span>DanhMucTacGia!$A$2:$E$81</span>`: lấy dữ liệu từ sheet tra cứu và cố định vùng.
* `<span>2</span>`: trả về tên tác giả ở cột thứ hai trong vùng tra cứu.
* `<span>FALSE</span>`: yêu cầu khớp chính xác mã tác giả.
* `<span>IFERROR</span>`: nếu tra cứu gặp lỗi, hiển thị `<span>Chưa </span><span>cập </span><span>nhật</span>`.

### 5. Công thức ô E2 và F2

Ô E2 – Phòng trình bày:

```
=IFERROR(VLOOKUP(TRIM(C2),DanhMucTacGia!$A$2:$E$81,4,FALSE),"Chưa cập nhật")
```

Ô F2 – Thù lao:

```
=IFERROR(VLOOKUP(TRIM(C2),DanhMucTacGia!$A$2:$E$81,5,FALSE),"Chưa cập nhật")
```

Chỉ cần thay `<span>Col_index_num</span>`vì bảng tra cứu đã có đủ thông tin:

* `<span>2</span>`: Tên tác giả.
* `<span>4</span>`: Phòng trình bày.
* `<span>5</span>`: Thù lao.

Các thành phần còn lại giữ nguyên để bảo đảm tra cứu đúng mã tác giả và xử lý lỗi.

### 6. Điều gì xảy ra nếu bỏ dấu `<span>$</span>`?

Nếu công thức D2 sử dụng vùng tra cứu không cố định:

```
=IFERROR(VLOOKUP(TRIM(C2),DanhMucTacGia!A2:E81,2,FALSE),"Chưa cập nhật")
```

Khi kéo công thức từ D2 xuống D5, Excel tự động dịch vùng tra cứu theo số dòng.


| Ôcôngthức | Vùngtracứu          |
| ------------ | --------------------- |
| D2           | `<span>A2:E81</span>` |
| D3           | `<span>A3:E82</span>` |
| D4           | `<span>A4:E83</span>` |
| D5           | `<span>A5:E84</span>` |

Hậu quả:Vùng tra cứu bị dịch xuống, khiến một số tác giả ở đầu danh sách không còn nằm trong vùng tìm kiếm và có thể dẫn đến kết quả sai hoặc báo `<span>Chưa </span><span>cập </span><span>nhật</span>`.

Cách khắc phục:Cố định vùng bằng `<span>$A$2:$E$81</span>`để tất cả công thức đều tra cứu trên cùng một bảng dữ liệu.

Kết luận:Kết hợp `<span>VLOOKUP</span>`, `<span>TRIM</span>`, `<span>IFERROR</span>`và dấu `<span>$</span>`giúp tự động điền tên tác giả, phòng trình bày và thù lao, đồng thời hạn chế lỗi khoảng trắng, mã không tồn tại và vùng tra cứu bị thay đổi.

Lưu ý:Nếu Excel trên máy bạn sử dụng dấu chấm phẩy để phân cách tham số, hãy thay các dấu phẩy `<span>,</span>`trong công thức bằng dấu chấm phẩy `<span>;</span>`.
