import mongoose from "mongoose";

export default async function connectToDB(id) {
    if (mongoose.connections[0].readyState) {
        await mongoose.connect(process.env.MONGO_URI)
    }
    else {
        return
    }
}