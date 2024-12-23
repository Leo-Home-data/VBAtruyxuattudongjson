# VBAtruyxuattudongjson
Chương trình được xây dựng từ VBA truy xuất json từ Monggo vào EXCEL
# Cách sử dụng:
Download file: '.xml'; 
Enable Macro trong Excel
Mở file và làm theo hướng dẫn sau:
1. Bạn lưu ý Mẫu JSON trong Chương trình này cài đặt sẵn dạng:
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

# Bạn phải vào Module "truyxuatjson" chỉnh lại mẫu JSON phù hợp với nhu cầu bên bạn
# Cũng trong Module này, bạn thay thế vị trí tệp json mà bạn lưu trữ ở dòng này: 
jsonText = jsonFile.OpenTextFile("C:\Users\pc\Desktop\DATA-SELF-RESEARCH\Data Analysis\NoSQL\VBA-JSON\VBA-TEST-1.json", 1).ReadAll

# Sau đó bạn cứ Nhấn Button "Rut tien" trong file '.xml'
Mọi thứ sẽ vận hành tự động.
