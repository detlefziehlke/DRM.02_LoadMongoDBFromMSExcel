"use strict";

var mongoose = require('mongoose');
var Schema = mongoose.Schema;

var tasksSchema = new Schema({
  subject: String,
  project: {
    type: Schema.ObjectId,
    ref: 'projects'
  }
});

mongoose.model('tasks', tasksSchema);
