import mongoose from "mongoose";

const certificateSchema = mongoose.Schema({
    name: {
        type: String,
        required: true,
    },
    Occasion: {
        type: String,
        required: true,
    },
    date: {
        type: String,
        required: true,
    },
    cid: {
        type: Number,
        required: true,
    },
});

mongoose.models = {}
module.exports = mongoose.model('certificates', certificateSchema)