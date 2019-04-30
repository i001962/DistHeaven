import * as React from 'react';

export interface HeaderProps {
    title: string;
}

export class Header extends React.Component<HeaderProps, any> {
    constructor(props, context) {
        super(props, context);
    }

    render() {
        return (
            <div id='content-header'>
                <div className='padding'>
                    <h1>{this.props.title}</h1>
                    <p>Select range eg C1:D1 with correlation coeffiecent in one cell and number of trials in the other</p>
                </div>
            </div>
        );
    }
}
