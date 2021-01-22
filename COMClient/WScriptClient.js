// Works as expected:
var server = new ActiveXObject("ComServerTlb.ServerTlb");
var pi = server.ComputePi();

WScript.Echo("PI: " + pi);