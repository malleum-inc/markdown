import * as React from 'react';
export interface FooterProps {}

export default class Footer extends React.Component<FooterProps> {
    render() {
        return (
            <div className='ms-font-s footer' style={{textAlign: 'right'}}>Copyright &copy; 2018 <a href="https://redcanari.com">Red Canari, Inc.</a></div>
        );
    }
}
