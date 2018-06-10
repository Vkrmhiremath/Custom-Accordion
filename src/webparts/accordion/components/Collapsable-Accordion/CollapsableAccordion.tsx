import * as React from 'react';
import styles from './CollapsableAccordion.module.scss';


export interface ICollapsableAccordionStates {
    open: boolean;
}

export interface ICollapsableAccordionProps {
    headerClass?: string;
    contentWrapperClass?: string;
    isOpen?: boolean;
    onToggle?: any;
    headerChildren?: any;
    title?: string;
};


export default class CollapsableAccordion extends React.Component<ICollapsableAccordionProps, ICollapsableAccordionStates>{

    public constructor(props) {
        super(props);
        this.state = {
            open: !!props.isOpen,
        };
    }

    public componentWillReceiveProps(nextProps) {
        if (this.props.isOpen !== nextProps.isOpen) {
            this.toggle();
        }
    }

    public componentDidUpdate(prevProps, prevState) {
        if (prevState.open !== this.state.open && this.props.onToggle) {
            this.props.onToggle(this.state.open);
        }
    }

    private toggle = () => {
        this.setState(prevState => {
            return {
                open: !prevState.open,
            };
        });
    };


    public render() {
        return (
            <div className={styles.CollapsableAccordion}>
                <div
                    onClick={this.toggle}
                    className={
                        this.state.open
                            ? this.props.headerClass ? styles.AccordionHeader + " " + this.props.headerClass : styles.AccordionHeader
                            : styles.AccordionHeader + " " + styles.AccordionHeaderCollapsed
                    }
                >
                    {this.props.title}
                    {this.props.headerChildren ? this.props.headerChildren(this.state.open) : null}
                </div>
                <div className={
                    this.state.open
                        ? this.props.contentWrapperClass ? styles.AccordionBody + " " + this.props.contentWrapperClass : styles.AccordionBody
                        : styles.AccordionBody + " " + styles.AccordionBodyCollapsed
                }>
                    {this.props.children}
                </div>
            </div>
        );
    }


}