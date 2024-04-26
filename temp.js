import { initializeApp } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-app.js";
import { getAuth,createUserWithEmailAndPassword, } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-app.js";


const firebaseConfig = {
  apiKey: "AIzaSyBpp9ZcMYgvNf-fRRroV4xjcVlAjTbKNTE",
  authDomain: "login-details-a2221.firebaseapp.com",
  databaseURL: "https://login-details-a2221-default-rtdb.firebaseio.com",
  projectId: "login-details-a2221",
  storageBucket: "login-details-a2221.appspot.com",
  messagingSenderId: "530026696513",
  appId: "1:530026696513:web:9003a9208c8d0e5de0dbc0"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth();

window.Register = function (e) {
  e.preventDefault();
  var email = document.getElementById("email").value;
  var userPassword = document.getElementById("userPassword").value;

  var obj = {
    email: email,
   };

  // Use createUserWithEmailAndPassword correctly and pass userPassword as the password
  createUserWithEmailAndPassword(auth, obj.email, obj.userPassword)
    .then(function (success) {
      alert("Successfully registered");
    })
    .catch(function (err) {
      alert("Error: " + err.message);
    });

  console.log(obj);
};
