"use strict";

var mongoose = require('mongoose');
var Schema = mongoose.Schema;

var projectsSchema = new Schema({
  name: String,
  area:String
});

mongoose.model('projects', projectsSchema);
