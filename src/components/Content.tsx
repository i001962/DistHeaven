import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export interface ContentProps {
    buttonLabel: string;
    click: any;
}

export class Content extends React.Component<ContentProps, any> {
    constructor(props, context) {
        super(props, context);
    }

    render() {
        return (
            <div id={this.props.buttonLabel}>
                <div className='padding'>
                    {/* <p>{this.props.message}</p> */}
                    <br />
                    <br />  <br />  <br />  <br />
                    <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                </div>
            </div>
        );
    }
}
