import React from 'react';
import * as XLSX from 'xlsx';

class AddToExcel extends React.Component {
  addToExcel = () => {
    // Charger le fichier Excel existant
    const fileReader = new FileReader();
    fileReader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Modifier le fichier Excel en ajoutant de nouvelles données
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const newData = [
        ['New Name', 'New Age'],
        ['Mike', 35],
        ['Emily', 28]
      ];
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      range.e.r += newData.length - 1;
      XLSX.utils.sheet_add_aoa(worksheet, newData, { origin: -1 });

      // Sauvegarder le fichier Excel mis à jour
      const updatedData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([updatedData], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'updated_data.xlsx');
      document.body.appendChild(link);
      link.click();
    };

    // Charger le fichier existant (vous devez avoir un input[type=file] dans votre JSX)
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.onchange = (event) => {
      const file = event.target.files[0];
      fileReader.readAsArrayBuffer(file);
    };
    input.click();
  };

  render() {
    return (
      <button onClick={this.addToExcel}>
        Add to Excel
      </button>
    );
  }
}

export default AddToExcel;
