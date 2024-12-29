# VBAtruyxuattudongjson

Chương trình được xây dựng từ VBA truy xuất json từ Monggo vào EXCEL. Mục tiêu để giúp quá trình truy xuất data quy mô vùa phải với thời gian nhanh, tự động hoá, tiết kiệm chi phí nhân sự và thời gian thao tác. Bằng cách vận dụng các hàm đơn giản và cài đặt các môi trường để kết nối làm việc với MongoDB như: MSXML và ADOBD

Sau đó chúng tôi, để đơn giản hoá, chúng tôi chọn phương thức Export truy xuất tự động JSON vào EXCEL theo phương thức: Export JSON từ Mongo và sau đó Import vào EXCEL.

Cách này đơn giản và phù hợp cho người mới bắt đầu làm quen truy xuất data phi cấu trúc bằng VBA.

Ngoài ra bạn còn có nhiều cách khác như: 

- sử dụng driver MongoDB ODB được cấp miễn phí,
            
- Sử dụng Python làm cầu nối nếu bạn chuyên về Python, code dưới đây là 1 ví dụ để bạn tham khảo. Tuy nhiên chúng ta cứ thoải mái vì File này chỉ cần bạn Export sẵn JSON là được, đơn giản và nhanh, tiện lợi.

                     ### Python Script (mongo_to_json.py):

            import json

                        from pymongo import MongoClient

                        def query_mongo():

                                    client = MongoClient("mongodb://localhost:27017/")*

                                    db = client["database_name"]**

                                    collection = db["collection_name"]***

                                    data = collection.find({}, {"_id": 0}) # Lấy tất cả dữ liệu, bỏ qua
                                    `_id`****

                                    with open("output.json", "w") as f:****
                                    json.dump(list(data), f)*

                        if __name__ == "__main__

                        query_mongo()*

            ### VBA Code để gọi Python:

            Sub RunPythonScript()

                        Dim shell As Object

                        Set shell = CreateObject("WScript.Shell")

                        ' Chạy Python script

                        shell.Run "python C:\path\to\mongo_to_json.py", 1, True

                        ' Xử lý dữ liệu sau khi script hoàn thành

                        MsgBox "Python script đã hoàn tất!"

            End Sub


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
