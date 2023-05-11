import React, { useEffect, useState } from 'react'
import { useRouter } from 'next/router'

const Slug = () => {
    const router = useRouter();
    const [auth, setAuth] = useState(false)
    const { slug } = router.query;
    useEffect(() => {
        if (router.isReady) { return }
        else {
            const verifyCertificate = async () => {
                console.log("Present Here", slug);
                const request = await fetch(`/api/contact`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(slug)
                })
                const response = await request.json();
                console.log(response);
                if (response.success) {
                    setAuth(true);
                }
                setAuth(true)
            }
            if (auth) {
                verifyCertificate();
            }
        }
    })
    return (
        <div style={{ display: "flex", justifyContent: "center", background: "url('/background.jpg')", width: '100%', height: '100vh' }}>
            {router.isReady && slug.length === 12 && <div style={{ fontSize: "30px", marginTop: "30px", color: "white" }}>Verified ✅</div>}
            {router.isReady && slug.length < 3 && <div style={{ fontSize: "30px", marginTop: "30px", color: "white" }}>Not Verified ❌</div>}
        </div>
    )
}
export default Slug