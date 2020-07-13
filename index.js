var express = require('express');
var path = require('path');
var fs = require("fs");
var app = express();

app.use('/assets',express.static(path.join(__dirname, 'assets')));

app.get('/', function (req, res) {
    res.sendFile("main.html",{root: __dirname });
});

var server = app.listen(3000);
