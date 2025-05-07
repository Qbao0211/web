<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Quản lý Item và Ảnh</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 40px;
      background: #f4f4f4;
    }
    h2 {
      text-align: center;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      background: #fff;
      margin-top: 20px;
    }
    th, td {
      border: 1px solid #ccc;
      padding: 10px;
    }
    th {
      background-color: #eee;
    }
    .image-dropzone {
      width: 120px;
      height: 100px;
      border: 2px dashed #999;
      text-align: center;
      line-height: 100px;
      color: #777;
      cursor: pointer;
      position: relative;
      overflow: hidden;
    }
    .image-dropzone img {
      max-width: 100%;
      max-height: 100%;
      position: absolute;
      top: 0;
      left: 0;
      object-fit: contain;
    }
    button {
      margin-top: 20px;
      padding: 10px 20px;
      font-size: 16px;
      cursor: pointer;
    }
  </style>
</head>
<body>

  <h2>Quản lý danh sách Item và ảnh</h2>

<div style="margin-bottom: 10px;">
  <label for="bulkInput">Dán danh sách Tên Item (mỗi dòng 1 item):</label><br>
  <textarea id="bulkInput" rows="5" style="width: 100%; margin-top: 5px;"></textarea><br>
  <button onclick="addItemsFromList()">Thêm từ danh sách</button>
</div>

<button onclick="addRow()">Thêm Item</button>
<button onclick="exportData()">Xuất Excel</button>

  <table id="itemTable">
    <thead>
      <tr>
        <th>STT</th>
        <th>Tên Item</th>
        <th>Mô Tả</th>
        <th>Ảnh</th>
      </tr>
    </thead>
    <tbody>
    </tbody>
  </table>

  <script>
    let count = 0;

    function addRow() {
      count++;
      const table = document.getElementById('itemTable').getElementsByTagName('tbody')[0];
      const row = table.insertRow();

      const cellSTT = row.insertCell(0);
      cellSTT.innerText = count;

      const cellName = row.insertCell(1);
      cellName.innerHTML = '<input type="text" name="name">';

      const cellDesc = row.insertCell(2);
      cellDesc.innerHTML = '<textarea name="desc" rows="2"></textarea>';

      const cellImage = row.insertCell(3);
      const dropzone = document.createElement('div');
      dropzone.className = 'image-dropzone';
      dropzone.innerText = 'Kéo & thả ảnh vào đây';
      dropzone.ondragover = (e) => e.preventDefault();
      dropzone.ondrop = (e) => handleDrop(e, dropzone);
      cellImage.appendChild(dropzone);
    }

  function addItemsFromList() {
    const input = document.getElementById('bulkInput').value.trim();
    if (!input) return alert("Vui lòng nhập danh sách tên item!");

    const lines = input.split("\n");
    lines.forEach(name => {
      name = name.trim();
      if (name) {
        addRowWithName(name);
      }
    });
    document.getElementById('bulkInput').value = ""; // Xóa sau khi thêm
  }

  function addRowWithName(name) {
    count++;
    const table = document.getElementById('itemTable').getElementsByTagName('tbody')[0];
    const row = table.insertRow();

    const cellSTT = row.insertCell(0);
    cellSTT.innerText = count;

    const cellName = row.insertCell(1);
    cellName.innerHTML = `<input type="text" name="name" value="${name}">`;

    const cellDesc = row.insertCell(2);
    cellDesc.innerHTML = '<textarea name="desc" rows="2"></textarea>';

    const cellImage = row.insertCell(3);
    const dropzone = document.createElement('div');
    dropzone.className = 'image-dropzone';
    dropzone.innerText = 'Kéo & thả ảnh vào đây';
    dropzone.ondragover = (e) => e.preventDefault();
    dropzone.ondrop = (e) => handleDrop(e, dropzone);
    cellImage.appendChild(dropzone);
  }

    function handleDrop(e, dropzone) {
      e.preventDefault();
      const file = e.dataTransfer.files[0];
      if (!file.type.startsWith('image/')) return;

      const reader = new FileReader();
      reader.onload = function(evt) {
        dropzone.innerHTML = `<img src="${evt.target.result}" class="preview-image">`;
        dropzone.dataset.image = evt.target.result;
      }
      reader.readAsDataURL(file);
    }

    function exportData() {
      const rows = document.querySelectorAll("#itemTable tbody tr");
      const data = [];

      rows.forEach(row => {
        const name = row.querySelector("input[name='name']").value;
        const desc = row.querySelector("textarea[name='desc']").value;
        const image = row.querySelector(".image-dropzone").dataset.image || "";
        data.push({ name, desc, image });
      });

      fetch("/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      })
      .then(response => response.blob())
      .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "danh_sach_item.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
      })
      .catch(err => {
        alert("Có lỗi xảy ra khi xuất Excel!");
        console.error(err);
      });
    }
  </script>
</body>
</html>

from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
import os, io, base64, uuid

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/export', methods=['POST'])
def export():
    data = request.json

    wb = Workbook()
    ws = wb.active
    ws.title = "Item Data"
    ws.append(['STT', 'Tên Item', 'Mô Tả', 'Ảnh'])

    temp_dir = 'temp_images'
    os.makedirs(temp_dir, exist_ok=True)

    for idx, item in enumerate(data, start=1):
        ws.append([idx, item['name'], item['desc']])

        if item['image']:
            image_data = item['image'].split(',')[1]
            image_bytes = base64.b64decode(image_data)

            image_path = os.path.join(temp_dir, f"{uuid.uuid4()}.png")
            with open(image_path, 'wb') as f:
                f.write(image_bytes)

            # Resize to fit Excel cell
            img = Image.open(image_path)
            img.thumbnail((100, 100))
            img.save(image_path)

            excel_img = ExcelImage(image_path)
            cell_ref = f'D{idx + 1}'
            ws.add_image(excel_img, cell_ref)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # Xoá ảnh tạm
    for f in os.listdir(temp_dir):
        os.remove(os.path.join(temp_dir, f))

    return send_file(output, as_attachment=True, download_name='danh_sach_item.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)
