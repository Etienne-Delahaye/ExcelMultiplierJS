<template>
  <div class="YOLO">
    <h1>{{ msg }}</h1>

    <div
      style="padding: 20px; margin: auto; border-radius: 3px; border-style: solid; border-color: black; width: 75%; text-align: left"
    >
      <ul>
        <li>Les entêtes doivent être sur la première ligne.</li>
        <li>Le même séparateur est utilisé pour toute les colonnes. Si besoin de séparer selon "," sur une colonne A, et selon " " pour une colonne B, il faut lancer le script deux fois.</li>
        <li>Attention à ne pas laisser trainer des espaces.</li>
        <li>
          Toute la ligne est dupliquée, y compris les cases avec des chiffres,
          <span style="color:red">attention au totaux</span> !
        </li>
        <li>
          Tout les calculs se font dans le navigateur, il n'y a aucun échange sur le réseau. <span style="color:orange">Si le fichier est gros, c'est votre ordinateur qui ralentit</span> :) <span style="color:green">Mais niveau confidentialité des données c'est safe.</span>
        </li>
      </ul>
    </div>

    <p>
      <input id="file" accept=".xls, .xlsx" type="file" @change="changeFile" />
    </p>

    <div v-if="sheets_name !== null">
      <label for="select_sheet" style="margin:10px">Sélectionner le tab</label>
      <select id="select_sheet" v-model="selected_sheet_name">
        <option disabled value>Please select one</option>
        <option v-for="sheet_name in sheets_name" :key="sheet_name">{{ sheet_name }}</option>
      </select>
    </div>

    <div style="color: red; margin:10px">{{ error_headers }}</div>
    <div style="color: red; margin:10px">{{ error_sep }}</div>
    <div style="color: red; margin:10px">{{ error_col }}</div>

    <div v-if="sheet !== null">
      <span>Cocher les colonnes à démultiplier:</span>
      <span v-for="col in headers" :key="col">
        <input
          type="checkbox"
          :id="col"
          :value="col"
          v-model="col_to_duplicate"
          style="margin:10px"
        />
        <label :for="col">{{ col }}</label>
      </span>
    </div>

    <div>
      <textarea style="white-space: pre-line" id="sep" v-model="sep" placeholder="Séparateur" />
    </div>

    <button v-if="this.sheet !== null" @click="go">Expand</button>

    <div v-if="content2 !== ''">
      <button @click="download">Download</button>
    </div>
  </div>
</template>

<script>
import XLSX from "xlsx";

export default {
  name: "HelloWorld",
  props: {
    msg: String,
  },

  data: function () {
    return {
      content: "",
      content2: "",
      excel_file: null,
      selected_sheet_name: null,
      col_to_duplicate: [],
      error_headers: "",
      error_sep: "",
      error_col: "",
      sep: "",
    };
  },

  computed: {
    sheets_name: {
      get: function () {
        if (this.excel_file === null) return null;
        return this.excel_file.SheetNames;
      },
    },
    sheet: {
      get: function () {
        if (this.selected_sheet_name === null) return null;
        return this.excel_file.Sheets[this.selected_sheet_name];
      },
    },

    headers: {
      get: function () {
        return this.get_headers();
      },
    },
  },

  methods: {
    get_headers: function () {
      if (this.sheet === null) return null;
      var headers = [];
      var range = XLSX.utils.decode_range(this.sheet["!ref"]);
      var C,
        R = range.s.r; /* start in the first row */
      /* walk every column in the range */
      for (C = range.s.c; C <= range.e.c; ++C) {
        var cell = this.sheet[
          XLSX.utils.encode_cell({ c: C, r: R })
        ]; /* find the cell in the first row */

        var hdr = "UNKNOWN " + C; // <-- replace with your desired default
        if (cell && cell.t) hdr = XLSX.utils.format_cell(cell);

        headers.push(hdr);
      }
      if (headers.length != new Set(headers).size) {
        this.error_headers =
          "There are columns with the same name resulting in conflict, please change this";
      } else {
        this.error_headers = "";
      }
      return headers;
    },

    changeFile: function (e) {
      (this.excel_file = null), (this.selected_sheet_name = null);
      this.col_to_duplicate = [];
      this.content = "";
      this.content2 = "";

      var f = e.target.files[0];
      var reader = new FileReader();
      reader.onload = (e) => {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        this.excel_file = workbook;
        this.selected_sheet_name = this.excel_file.SheetNames[0];
      };
      reader.readAsArrayBuffer(f);
    },

    go: function () {
      this.content2 = "";
      if (this.col_to_duplicate.length < 1) {
        this.error_col = "Veuillez indiquer au moins une colonne à dupliquer";
      } else this.error_col = "";
      if (this.sep === null || this.sep === "") {
        this.error_sep =
          'Veuillez indiquer un séparateur (par exemple: ",", " ", <enter>...)';
      } else this.error_sep = "";

      if (this.error_headers || this.error_sep || this.error_col) return;

      var data = XLSX.utils.sheet_to_json(this.sheet, { header: 1 });
      this.content = XLSX.utils.sheet_to_json(this.sheet, { header: 1 }); // just a save

      for (const col of this.col_to_duplicate) {
        // for every col_to_duplicate, we iter over the dataset
        var result = [];
        const index = this.headers.findIndex((e) => e == col);

        data.map((row) => {
          // for each line
          if (
            row[index] == null ||
            row[index].toString().split(this.sep).length <= 1
          ) {
            result.push(row.slice()); // to duplicate !
          } else {
            for (const dup of row[index].toString().split(this.sep)) {
              if (dup.trim().length > 0) {
                result.push(row.slice()); // to duplicate
                result[result.length - 1][index] = dup;
              }
            }
          }
        }); // end iter over row
        data = result;
        result = [];
      } //end col_to_duplicate

      this.content2 = data;
    },

    download: function () {
      var wb = XLSX.utils.book_new();
      var ws = XLSX.utils.aoa_to_sheet(this.content2);
      XLSX.utils.book_append_sheet(
        wb,
        ws,
        this.selected_sheet_name + "_deduplicated"
      );
      XLSX.writeFile(wb, "output.xlsx");
    },
  },
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
</style>
