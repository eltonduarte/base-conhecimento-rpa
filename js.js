# Seleciona menu com js
document.querySelector("#ddlPorto").options[6].selected = "SELECTED";
document.querySelector("#ddlPorto").dispatchEvent(new Event('change'));

