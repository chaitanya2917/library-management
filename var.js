<script type="module">
  import {io} from "https://cdn.socket.io/4.3.2/socket.io.esm.min.js";
</script>;

const socket = io();
let data;

socket.on("telemetry", (data) => {
  data = data;
});
