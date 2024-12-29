# VBAtruyxuattudongjson

Chương trình được xây dựng từ VBA truy xuất json từ Monggo vào EXCEL

## Cách sử dụng:
1 - Download file: '.xml'; 
2 - Enable Macro trong Excel
3 - Cài đặt file JsonConverter vào VBA để xử lý được dữ liệu json
4 - Mở file và làm theo hướng dẫn sau:

### Bạn lưu ý Mẫu JSON trong Chương trình này cài đặt sẵn dạng:

   {

  "_id": {

    "$oid": "6763a2db35ebb992664a8e01"

  },

  "transaction_id": "TRX12345",

  "customer": {

    "name": "Nguyen Van A",
 
    "email": "nguyenvana@example.com",

    "ma_khach_hang": "001"

  },

  "amount": 15000000,

  "currency": "VND",

  "Loai_dich_vu": "rut tien",

  " Trang thai giao dich": "Thanh cong"

}

## Module "truyxuatjson"

1- Bạn phải vào Module "truyxuatjson" chỉnh lại mẫu JSON phù hợp với nhu cầu bên bạn

2 - Cũng trong Module này, bạn thay thế vị trí tệp json mà bạn lưu trữ ở dòng này: 

jsonText = jsonFile.OpenTextFile("C:\Users\pc\Desktop\DATA-SELF-RESEARCH\Data Analysis\NoSQL\VBA-JSON\VBA-TEST-1.json", 1).ReadAll

## Sau đó bạn cứ Nhấn Button "Rut tien" trong file '.xml'
Mọi thứ sẽ vận hành tự động.

## Chúc bạn thành công
### Tác giả: DA Lý Tú Anh - Domain Sale-Marketing - Quản lý bán hàng kênh MT cho Uniclever 2024
