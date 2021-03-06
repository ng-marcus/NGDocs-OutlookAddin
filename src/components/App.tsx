import * as React from 'react';
// import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import Progress from './Progress';



export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
    attachments: Office.AttachmentDetails[];
}

export interface AppState {
    selectedAttachments: string[];
    attachments: Office.AttachmentDetails[];
    projects: string[];
}



export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            selectedAttachments: [],
            attachments: props.attachments,
            projects: []
        };
    }

    static getDerivedStateFromProps(props, state) {
        console.log('derive state from props');

        if (props.attachments.length !== state.attachments.length) {
            console.log(props);
            let selected = [];
            props.attachments.forEach(a => { if (!a.isInline) { selected.push(a.id); } });
            return { selectedAttachments: selected, attachments: props.attachments };
        } else {
            return state;
        }
    }


    componentDidMount() {
        console.log('Component did mount');


        this.handleInputChange = this.handleInputChange.bind(this);
        this.handleSubmit = this.handleSubmit.bind(this);
        this.attachmentTokenCallback = this.attachmentTokenCallback.bind(this);

        console.log('fetch projects');
        fetch('https://ngdocs-spsupport.azurewebsites.net/api/GetProjectList?code=Oxv466R06RaWWswbayxdlru5KdMRS2Jm1/fg9XBeFyjoYabF8tGi0A==',
            {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                }
            })
            .then(response => response.json())
            .then(projects => {
                console.log(projects);

                this.setState({ ...this.state, projects: projects.projects });
            });





    }

    click = async () => {
        /**
         * Insert your Outlook code here
         */
    }

    attachmentTokenCallback(asyncResult) {
        if (asyncResult.status === 'succeeded') {

            console.log('token received');
            console.log(asyncResult.value);
            console.log(this.state.selectedAttachments);
            this.state.selectedAttachments.forEach(attachment =>
                fetch('https://prod-09.uksouth.logic.azure.com:443/workflows/0486e4ac7ec64370a4df4dde8feecb9d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BRNQBQ8ojyuwTe9AWXFEJgp_tfr_2YN-05M2VYJp8vE',
                    {
                        method: 'post',
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            from: 'm.roberts@nittygritty.net',
                            token: asyncResult.value,
                            attachment_id: attachment
                        })
                    })
                    .then(_response => console.log('sent')));

        } else {
            console.log(asyncResult.error.message);
        }
    }

    handleInputChange(event) {
        console.log(event);
        console.log(this.state.selectedAttachments);

        const target = event.target;


        if (target.checked) {
            console.log('add');
            console.log(target.value);
            let new_attachments = this.state.selectedAttachments;
            new_attachments.push(target.value);
            this.setState({ selectedAttachments: new_attachments });
        } else {
            console.log('remove');
            console.log(target.value);
            let new_attachments = this.state.selectedAttachments.filter(i => i !== target.value);
            this.setState({ selectedAttachments: new_attachments });

        }

        console.log(this.state.selectedAttachments);



    }

    handleSubmit(event) {
        console.log('File attachments');
        event.preventDefault();
        Office.context.mailbox.getCallbackTokenAsync(this.attachmentTokenCallback);

    }




    render() {
        const {
            title,
            isOfficeInitialized,
            attachments
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/logo-filled.png'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        const listAttachments = attachments.map((item, index) => (
            <li className='ms-ListItem' key={index}>
                <span className='ms-font-m ms-fontColor-neutralPrimary'>{item.name}</span>
                <input
                    name='toFile'
                    type='checkbox'
                    value={item.id}
                    checked={this.state.selectedAttachments.includes(item.id)}
                    onChange={this.handleInputChange} />
            </li>
        ));
        const listProjects = this.state.projects.map((item, index) => (
            <option key={'p' + index} value={item}>
                {item}
            </option>
        ));
        return (
            <div className='ms-welcome'>
                <Header logo='assets/logo-filled.png' title={this.props.title} message='File Attachments' />
                <form onSubmit={this.handleSubmit}>
                    <ul className='ms-List ms-welcome__features ms-u-slideUpIn10'>
                        {listAttachments}
                    </ul>
                    <br />
                    Select project:&nbsp;
                    <select >
                        {listProjects}
                    </select>
                    <br />
                    <br />
                    <input type='submit' value='File attachments' />
                </form>
            </div >
        );
    }
}
