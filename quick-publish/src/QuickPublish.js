import React, { useEffect, useState } from 'react';
import {connect, PublishAndNotify, quickpublish} from './sharepoint.js';

const QuickPublish = () => {
    const [data, setData] = useState(null);

    useEffect(() => {
        const fetchData = async () => {
            try {
                await connect(async () => {
                    try {
                        const response  = await PublishAndNotify();
                        console.log(response);
                        setData(response);
                    } catch (e) {
                        console.error(e);
                    }
                });
            } catch (error) {
                console.error('Error in async function:', error);
            }
        };

        fetchData();
    }, []);

    return (
        <div>
            {data ? (
                <div>Data: {data}</div>
            ) : (
                <div>Loading...</div>
            )}
        </div>
    );
};

export default QuickPublish;