const btn=document.querySelectorAll('a');
const body=document.querySelectorAll('section');
console.log(btn);
let temp;
let store;
let hideSection;

btn.forEach(function (buttons) {
   console.log(buttons);

   buttons.addEventListener('click',function(e){
     for(let i=1;i<=4;i++)
     {
        let a= document.querySelector(`li:nth-child(${i})`).innerText;
        console.log(a);
        hideSection=document.getElementById(`${a}`);
        console.log(hideSection);
        hideSection.style.display="none";
        console.log(a);
     }//to Hide All Section 

   temp1=e.target.innerText;
   store=document.getElementById(`${temp1}`);//For Show Current Page 

    if(temp1=='Register')
        store.style.display="flex";
    else 
        store.style.display="block";

   })//lose Of Event Listner

});



/*
let i;
const homeBtn=document.querySelector('li:nth-child(1)')
const registerBtn=document.querySelector('li:nth-child(2)')
const aboutusBtn=document.querySelector('li:nth-child(3)')
const contactusBtn=document.querySelector('li:nth-child(4)')

const homePage=document.getElementById("home")
const registerPage=document.getElementById("Register")
const aboutusPage=document.getElementById("About_us")
const contactusPage=document.getElementById("Contact_us")


homeBtn.addEventListener('click', function () {

    for(let i=1;i<=4;i++)
    {
        document.querySelectorAll(`li:nth-child(${i})`).parentid.style.display="none"
    }
    
    
    aboutusPage.style.display = 'none';
    contactusPage.style.display = 'none';
    registerPage.style.display = 'none';
    homePage.style.display='block';
      
});

registerBtn.addEventListener('click', function () {
    homePage.style.display = 'none';
    contactusPage.style.display = 'none';
    aboutusPage.style.display = 'none';
    registerPage.style.display='flex';
    
});

aboutusBtn.addEventListener('click', function () {
    homePage.style.display = 'none';
    contactusPage.style.display = 'none';
    registerPage.style.display = 'none';
    aboutusPage.style.display='block';
    
});

contactusBtn.addEventListener('click', function () {
    homePage.style.display = 'none';
    aboutusPage.style.display = 'none';
    registerPage.style.display = 'none';
    
    contactusPage.style.display='block';
    
});
*/

/*
<!DOCTYPE html>
<html>
<head>
    <title>HTML Form to Excel</title>
</head>
<body>
    <h2>Enter Data</h2>
    <form id="myForm">
        <label for="name">Name:</label>
        <input type="text" id="name" name="name"><br><br>
        
        <label for="age">Age:</label>
        <input type="number" id="age" name="age"><br><br>

        <button type="button" onclick="addData()">Add Data</button>
    </form>

    <h2>Export Data</h2>
    <button type="button" onclick="exportToExcel()">Export to Excel</button>
</br>

*/
        var data = [];

        function addData() {
            var name = document.getElementById("name").value;
            var Email = document.getElementById("email").value;
            var mob1 = document.getElementById("mob1").value;
            var mob2 = document.getElementById("mob2").value;
            var marks = document.getElementById("marks").value;
            var address = document.getElementById("address").value;
            var Description = document.getElementById("Description").value;
            data.push({ "Name": name, "Email": Email ,"mob1": mob1,"mob2": mob2,"marks":marks,"address":address,"Description":Description});
          //  document.getElementById("myForm").reset();
        }

        function exportToExcel() {
            if (data.length === 0) {
                alert("No data to export.");
                return;
            }

            var workbook = XLSX.utils.book_new();
            var worksheet = XLSX.utils.json_to_sheet(data);
            XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

            XLSX.writeFile(workbook, "test.xlsx");
        }

//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
