<script type="module"></script>
const firebaseConfig = {
    apiKey: "AIzaSyBpp9ZcMYgvNf-fRRroV4xjcVlAjTbKNTE",
    authDomain: "login-details-a2221.firebaseapp.com",
    projectId: "login-details-a2221",
    storageBucket: "login-details-a2221.appspot.com",
    messagingSenderId: "530026696513",
    appId: "1:530026696513:web:8b4e7b11462d8a0ce0dbc0"
  };

  const app = initializeApp(firebaseConfig);
  
  const auth=firebase.auth()
  const database=firebase.database()
  function register(){
    var email=document.getElementById("email").value;
    var fullName=document.getElementById("fullName").value;
    var userId = document.getElementById("userId").value;
    var userPassword = document.getElementById("userPassword").value;
    var confirmUserPassword=document.getElementById("confirmUserPassword").value;
    console.log(email);
    if(validate_email(email)==false){
        alert("Someasdfasdfthing is wrong")
        return
    }
    else{
        alert("something is right")
    }
    auth.createUserWithEmailAndPassword(email, userPassword)
    .then(function() {
        var user=auth.currentUser
        var database_ref=database.ref()
        var user_data={
            email: email,
            fullName: fullName, 
            userId: userId,
            last_login: Date.now()  
        }
        database_ref.child('users/'+ user.uid).set(user_data)
        console.log("Hello")
        alert("User Created!!")
        })
        
    .catch(function(error){
        var error_code=error.code
        var error_message=error.message
        alert(error_message)
    })
  }

  function validate_email(email){
    expression=/^[^@]+@\w+(\.\w+)+\w$/
    if(expression.test(email)==true){
        return true
    }
    else{
        return false
    }
  }


//     if(userId=="" || userPassword=="" || confirmUserPassword=="" || email==""){
//         alert("Field cannot be empty!!");
//         return false;
//     }
//     else if(userPassword!=confirmUserPassword){
//         alert("Password doesn't Match!!");
//         return false;
//     }
//     else{
//         alert("Registerd Successfully");
//     }
//     return true;
// }