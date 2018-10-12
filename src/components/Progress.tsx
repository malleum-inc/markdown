/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import * as React from 'react';
import {Spinner, SpinnerSize} from 'office-ui-fabric-react/lib/components/Spinner';

export interface ProgressProps {
    logo: string;
    message: string;
    title: string;
}

export default class Progress extends React.Component<ProgressProps> {
    render() {
        const {
            logo,
            title,
            message,
        } = this.props;

        return (
            <>
                <div style={{flex: 1}}/>
                <div style={{border: 'none'}} className='main ms-md12 ms-lg12 ms-fontColor-white'>
                    <div className='loadingSplash'>
                        <img style={{width: '20%'}} src={logo} alt={title} title={title}/>
                        <Spinner size={SpinnerSize.large} label={message}/>
                    </div>
                </div>
                <div style={{flex: 1}}/>
            </>
        );
    }
}
