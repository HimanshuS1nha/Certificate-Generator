import React, { useEffect, useState } from 'react'
import { useRouter } from 'next/router'

const Id = () => {
    const router = useRouter();
    const { id } = router.query;
    const [auth, setAuth] = useState(false);
    useEffect(() => {
        if (router.isReady) { return }
        else {
            const verifyCertificate = async () => {
                console.log("Present Here", router);
                const request = await fetch(`${process.env.NEXT_PUBLIC_HOST}/api/contact`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ id })
                })
                const response = await request.json();
                console.log(response);
                setAuth(true)
            }
            verifyCertificate();
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [router.isReady])

    return (
        <>
            {auth && <div>VERIFIED</div>}
        </>
    )
}

export default Id