// Works as expected:
var server = new ActiveXObject("ComServerVbs.ServerVbs");
var pi = server.ComputePi();

WScript.Echo("PI: " + pi);