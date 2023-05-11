// Next.js API route support: https://nextjs.org/docs/api-routes/introduction
import connectToDB from "../../../MongoDB/dbConnect";
import certificates from '../../../models/certificates'

export default async function handler(req, res) {
  if (req.method === 'POST') {
    const id = req.params;
    console.log(id);
    await connectToDB();
    try {
      const certificate = await certificates.findOne({ cid: id });
      if (certificate) {
        res.status(200).json({ success: true });
      }
      else {
        res.status(200).json({ success: false });
      }
    } catch (error) {
      res.status(500).json({ success: false });
      console.log(error);
    }
  }
}