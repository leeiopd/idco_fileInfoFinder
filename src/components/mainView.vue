<template>
  <v-card class="ma-5">
    <v-card-title>
      <v-row align="center" justify="center">
        <v-col cols="2">
          <v-btn samll color="primary" class="btn btn-info" @click="onPickFile">
            엑셀 파일
          </v-btn>
          <input
            type="file"
            style="display: none"
            ref="fileInput"
            accept=".xlsx"
            @change="onFilePicked"
          />
        </v-col>
        <v-col cols="6" justify="center" align="center">
          <v-row justify="center" align="center">
            <v-text-field
              clearable
              v-model="search"
              label="주소 검색"
              @keyup.enter="addressSearch"
            >
            </v-text-field>
            <v-btn icon color="primary" @click="addressSearch">
              <v-icon>mdi-magnify</v-icon>
            </v-btn>
          </v-row>
        </v-col>
        <v-col cols="1">
          <v-btn icon color="green" @click="fields = []">
            <v-icon>mdi-refresh</v-icon>
          </v-btn>
        </v-col>
        <v-col cols="1">
          <v-btn icon color="indigo" @click="exportExcel">
            <v-icon>mdi-download</v-icon>
          </v-btn>
        </v-col>
      </v-row>
    </v-card-title>
    <v-card-text>
      <v-data-table
        :headers="headers"
        :items="fields"
        :items-per-page="-1"
        :hide-default-footer="true"
        class="elevation-1"
      >
        <template v-slot:item.address="{ item }">
          <a
            target="_blank"
            :href="
              `https://map.kakao.com/?map_type=TYPE_MAP&q=` +
                item.address +
                `&urlLevel=3`
            "
            >{{ item.address }}</a
          >
        </template>
        <template v-slot:no-data>
          검색 결과가 없습니다.
        </template>
      </v-data-table>
    </v-card-text>
    <v-dialog persistent v-model="loadingDialog" max-width="400">
      <v-card>
        <v-card-title class="mb-5">
          <v-row justify="center">
            <span>
              검색중...
            </span>
          </v-row>
        </v-card-title>
        <v-card-text class="mt-5">
          <v-row justify="center">
            <v-progress-circular
              :size="70"
              :width="7"
              color="primary"
              indeterminate
            ></v-progress-circular>
          </v-row>
        </v-card-text>
      </v-card>
    </v-dialog>
  </v-card>
</template>

<script>
import axios from "axios";
import XLSX from "xlsx";
export default {
  data() {
    return {
      search: "",
      fields: [],
      loadingDialog: false,
      headers: [
        {
          text: "주소",
          align: "center",
          sortable: false,
          value: "address",
        },
        {
          text: "면적",
          align: "center",
          sortable: false,
          value: "size",
        },
      ],
      file: null,
    };
  },
  methods: {
    onPickFile() {
      this.$refs.fileInput.click();
    },
    onFilePicked(event) {
      this.file = event.target.files[0];
    },
    async addressSearch() {
      if (!this.search) {
        return alert("검색어를 입력해 주세요!");
      }
      this.loadingDialog = true;
      let url =
        "http://api.vworld.kr/req/search?search=search&version=2.0&request=search&key=" +
        process.env.VUE_APP_API_VWORLD +
        "&format=json&query=" +
        this.search +
        "&type=ADDRESS&category=PARCEL";

      let status = true;

      const res = await axios
        .get(url)
        .then((result) => {
          return result.data.response;
        })
        .catch(() => {
          status = false;
          return;
        });

      if (!status) {
        return alert("인터넷 연결을 확인해 주세요");
      }

      if (res.status === "NOT_FOUND") {
        return alert("검색결과가 없습니다.\n주소를 확인 해 주세요.");
      }
      const fields = res.result.items;

      fields.forEach(async (field) => {
        var url =
          "http://apis.data.go.kr/1611000/nsdi/LandMoveService/attr/getLandMoveAttr?ServiceKey=" +
          process.env.VUE_APP_APIS_DATA +
          "&format=JSON&pnu=" +
          field.id;

        var xhr = new XMLHttpRequest();
        var comp = this;

        xhr.open("GET", url, false);

        xhr.onreadystatechange = function() {
          if (this.readyState === XMLHttpRequest.DONE) {
            if (this.status === 200) {
              comp.customerArray = JSON.parse(this.responseText);
            }
          }
        };

        xhr.send();

        const size =
          comp.customerArray.landMoves.field[
            comp.customerArray.landMoves.field.length - 1
          ].lndpclAr;

        const address = field.address.parcel;
        delete field.address;
        field.address = address;
        field.size = size;
      });

      this.fields = fields;
      this.loadingDialog = false;
    },

    async onChange(event) {
      if (!event.target.files) {
        return alert("excel 파일을 선택해 주세요!");
      }
      document.getElementById("fileUpload").click();
    },
    async readFile() {
      // eslint-disable-next-line no-unused-vars
      const fileName = this.file.name;

      // declare FileReader, temp result
      const reader = new FileReader();

      let tmpResult = {};
      // if you use "this", don't use "function(e) {...}"

      reader.onload = async (e) => {
        this.loadingDialog = true;
        let data = e.target.result;
        data = new Uint8Array(data);
        // get excel file
        // eslint-disable-next-line no-undef
        let excelFile = XLSX.read(data, { type: "array" });
        // get prased object
        excelFile.SheetNames.forEach(function(sheetName) {
          // eslint-disable-next-line no-undef
          const roa = XLSX.utils.sheet_to_json(excelFile.Sheets[sheetName], {
            header: 1,
          });
          if (roa.length) tmpResult[sheetName] = roa;
        });
        const sheet = tmpResult.Sheet1;

        const addressList = [];

        let check = false;

        for (let i = 0; i < 10; i++) {
          if (sheet[0][i] === "주소") {
            check = true;
          }
        }
        if (!check) {
          this.loadingDialog = false;
          return alert("주소 필드가 없습니다.");
        }
        for (let i = 1; i < sheet.length; i++) {
          for (let ii = 0; ii < 10; ii++) {
            if (sheet[0][ii] === "주소") {
              addressList.push(sheet[i][ii]);
            }
          }
        }

        const fields = [];

        async function getPnu(address) {
          let url =
            "http://api.vworld.kr/req/search?search=search&version=2.0&request=search&key=" +
            process.env.VUE_APP_API_VWORLD +
            "&format=json&query=" +
            address +
            "&type=ADDRESS&category=PARCEL";

          const res = await axios.get(url).then((result) => {
            return result.data.response;
          });

          if (res.status !== "NOT_FOUND") {
            for (let i = 0; i < res.result.items.length; i++) {
              fields.push(res.result.items[i]);
            }
          }
        }

        const promises = addressList.map(getPnu);
        // wait until all promises are resolved
        await Promise.all(promises);

        fields.forEach(async (field) => {
          var url =
            "http://apis.data.go.kr/1611000/nsdi/LandMoveService/attr/getLandMoveAttr?ServiceKey=" +
            process.env.VUE_APP_APIS_DATA +
            "&format=JSON&pnu=" +
            field.id;

          var xhr = new XMLHttpRequest();
          var comp = this;

          xhr.open("GET", url, false);

          xhr.onreadystatechange = function() {
            if (this.readyState === XMLHttpRequest.DONE) {
              if (this.status === 200) {
                comp.customerArray = JSON.parse(this.responseText);
              }
            }
          };

          xhr.send();

          const size =
            comp.customerArray.landMoves.field[
              comp.customerArray.landMoves.field.length - 1
            ].lndpclAr;

          const address = field.address.parcel;
          delete field.address;
          field.address = address;
          field.size = size;
        });

        this.fields = fields;
        this.loadingDialog = false;
      };
      reader.readAsArrayBuffer(this.file);
    },
    exportExcel() {
      if (this.fields.length === 0) {
        return alert("검색된 주소가 없습니다!");
      }
      const infos = [];

      for (let i = 0; i < this.fields.length; i++) {
        const tmp = { 주소: this.fields[i].address, 면적: this.fields[i].size };
        infos.push(tmp);
      }

      var infosWS = XLSX.utils.json_to_sheet(infos);

      var wb = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(wb, infosWS, "Sheet1");

      XLSX.writeFile(wb, "필지정보_검색결과.xlsx");
    },
  },
  watch: {
    file() {
      if (!this.file) {
        return;
      }
      this.readFile();
    },
  },
};
</script>

<style></style>
