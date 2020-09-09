<template>
  <v-app>
    <v-app-bar app color="blue" dark>
      <v-col justify="center">
        <v-row>
          <div class="d-flex align-center">
            <v-img
              alt="IDCo Logo"
              class="shrink mx-2"
              contain
              src="./assets/IDCo-logo.png"
              transition="scale-transition"
              width="60"
            />
          </div>
          <div class="ml-2 mt-1">
            토지 정보 검색기
          </div>
        </v-row>
      </v-col>
    </v-app-bar>

    <v-main>
      <mainView />
    </v-main>
  </v-app>
</template>

<script>
import mainView from "./components/mainView";

export default {
  name: "App",

  components: {
    mainView,
  },

  data: () => ({
    //
  }),
};
</script>
