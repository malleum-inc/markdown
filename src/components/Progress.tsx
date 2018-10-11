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
        } = this.props;

        return (
            <>
                <div className='header'/>
                <div className='main ms-font-xxl ms-bgColor-black ms-fontColor-white'>
                    <div style={{position: 'relative', top: '30%', height: 20, textAlign: 'center'}}>
                        <img style={{width: '20%'}} src={logo} alt={title} title={title}/>
                        <Spinner size={SpinnerSize.large} />
                    </div>
                </div>
                <div className='footer'/>
            </>
        );
    }
}
