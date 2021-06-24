// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Loader, Button, Text, List, Image, Flex } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";

import './sendConfirmationTaskModule.scss';
import { getDraftNotification, getConsentSummaries, sendDraftNotification } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: number;
    reactions?: number;
    responses?: number;
    succeeded?: number;
    failed?: number;
    throttled?: number;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
}

export interface SendConfirmationTaskModuleProps extends RouteComponentProps, WithTranslation {
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
    teamNames: string[];
    rosterNames: string[];
    groupNames: string[];
    allUsers: boolean;
    messageId: number;
}

class SendConfirmationTaskModule extends React.Component<SendConfirmationTaskModuleProps, IStatusState> {
    readonly localize: TFunction;
    private initMessage = {
        id: "",
        title: ""
    };

    private card: any;

    constructor(props: SendConfirmationTaskModuleProps) {
        super(props);
        this.localize = this.props.t;
        this.card = getInitAdaptiveCard(this.localize);

        this.state = {
            message: this.initMessage,
            loader: true,
            teamNames: [],
            rosterNames: [],
            groupNames: [],
            allUsers: false,
            messageId: 0,
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();

        let params = this.props.match.params;

        if ('id' in params) {
            let id = params['id'];
            this.getItem(id).then(() => {
                getConsentSummaries(id).then((response) => {
                    this.setState({
                        teamNames: response.data.teamNames.sort(),
                        rosterNames: response.data.rosterNames.sort(),
                        groupNames: response.data.groupNames.sort(),
                        allUsers: response.data.allUsers,
                        messageId: id,
                    }, () => {
                        this.setState({
                            loader: false
                        }, () => {
                            setCardTitle(this.card, this.state.message.title);
                            setCardImageLink(this.card, this.state.message.imageLink);
                            //setCardSummary(this.card, this.state.message.summary);
                            console.log("summaryDB : " + this.state.message.summary);
                            this.onContentStateChange(this.state.message.summary);
                            setCardAuthor(this.card, this.state.message.author);
                            if (this.state.message.buttonTitle && this.state.message.buttonLink) {
                                setCardBtn(this.card, this.state.message.buttonTitle, this.state.message.buttonLink);
                            }

                            let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                            adaptiveCard.parse(this.card);
                            let renderedCard = adaptiveCard.render();
                            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                            if (this.state.message.buttonLink) {
                                let link = this.state.message.buttonLink;
                                adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                            }
                        });
                    });
                });
            });
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            this.setState({
                message: response.data
            });
        } catch (error) {
            return error;
        }
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            return (
                <div className="taskModule">
                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                        <Flex className="scrollableContent" gap="gap.small">
                            <Flex.Item size="size.half">
                                <Flex column className="formContentContainer">
                                    <h3>{this.localize("ConfirmToSend")}</h3>
                                    <span>{this.localize("SendToRecipientsLabel")}</span>

                                    <div className="results">
                                        {this.renderAudienceSelection()}
                                    </div>
                                </Flex>
                            </Flex.Item>
                            <Flex.Item size="size.half">
                                <div className="adaptiveCardContainer">
                                </div>
                            </Flex.Item>
                        </Flex>
                        <Flex className="footerContainer" vAlign="end" hAlign="end">
                            <Flex className="buttonContainer" gap="gap.small">
                                <Flex.Item push>
                                    <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("PreparingMessageLabel")} labelPosition="end" />
                                </Flex.Item>
                                <Button content={this.localize("Send")} id="sendBtn" onClick={this.onSendMessage} primary />
                            </Flex>
                        </Flex>
                    </Flex>
                </div>
            );
        }
    }

    private onSendMessage = () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");
        sendDraftNotification(this.state.message).then(() => {
            microsoftTeams.tasks.submitTask();
        });
    }

    private getItemList = (items: string[]) => {
        let resultedTeams: IListItem[] = [];
        if (items) {
            resultedTeams = items.map((element) => {
                const resultedTeam: IListItem = {
                    header: element,
                    media: <Image src={ImageUtil.makeInitialImage(element)} avatar />
                }
                return resultedTeam;
            });
        }
        return resultedTeams;
    }

    private renderAudienceSelection = () => {
        if (this.state.teamNames && this.state.teamNames.length > 0) {
            return (
                <div key="teamNames"> <span className="label">{this.localize("TeamsLabel")}</span>
                    <List items={this.getItemList(this.state.teamNames)} />
                </div>
            );
        } else if (this.state.rosterNames && this.state.rosterNames.length > 0) {
            return (
                <div key="rosterNames"> <span className="label">{this.localize("TeamsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.rosterNames)} />
                </div>);
        } else if (this.state.groupNames && this.state.groupNames.length > 0) {
            return (
                <div key="groupNames" > <span className="label">{this.localize("GroupsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.groupNames)} />
                </div>);
        } else if (this.state.allUsers) {
            return (
                <div key="allUsers">
                    <span className="label">{this.localize("AllUsersLabel")}</span>
                    <div className="noteText">
                        <Text error content={this.localize("SendToAllUsersNote")} />
                    </div>
                </div>);
        } else {
            return (<div></div>);
        }
    }

    private onContentStateChange = (summaryDB) => {
        var jsonArr = [];
        var data;
        var count = 0;
        var summary = JSON.parse(summaryDB);
        console.log("summary : " + summary);

        for (let i = 0; i < summary.blocks.length; i++) {
            var element = summary.blocks[i];
            if(summary.blocks[i].text){
                var eText = "";
                eText = element.text + '\n'
               if(summary.blocks[i].type === "ordered-list-item"){
                    count = count + 1;
                    eText = count + ". " + element.text + '\n'
               } else if(summary.blocks[i].type === "unordered-list-item"){
                    eText = "* " + element.text + '\n'
               }
               var color = "default";
               var fontType = "default";
               var highlight = false;
               var isSubtle = false;
               var italic = false;
               var size = "default";
               var strikethrough = false;
               var underline = false;
               var weight = "default";
               var cAlign = "left";

               for(let j = 0; j < element.inlineStyleRanges.length; j++){
                    var sInline = element.inlineStyleRanges[j];
                    if(sInline){
                        if(sInline.style === "BOLD"){
                            weight = "bolder"
                        }
                        if(sInline.style === "ITALIC"){
                            italic = true
                        }
                        if(sInline.style === "UNDERLINE"){
                            underline = true
                        }
                        if(sInline.style === "STRIKETHROUGH"){
                            strikethrough = true
                        }
                        if(sInline.style.substring(0, 4) === "font"){
                            if(sInline.style.substring(0, 5) === "fontf"){
                                fontType = "monospace"
                            }else if(sInline.style.substring(0, 5) === "fonts"){
                                var fontSize = sInline.style.split('-');
                                if(fontSize[1] < 12){
                                    size = "small";
                                } else if(fontSize[1] < 18){
                                    size = "medium";
                                } else if(fontSize[1] < 48){
                                    size = "large";
                                } else if(fontSize[1] < 100){
                                    size = "extraLarge";
                                }
                            }
                        }

                        if(sInline.style.substring(0, 4) === "bgco"){
                            highlight = true;
                        }

                        if(sInline.style.substring(0, 4) === "colo"){
                            var rgb = sInline.style;
                            if(rgb === 'color-rgb(65,168,95)')
                                color = "good";
                            if(rgb === 'color-rgb(239,239,239)')
                                color = "light";
                            if(rgb === 'color-rgb(41,105,176)')
                                color = "accent";
                            if(rgb === 'color-rgb(243,121,52)')
                                color = "warning";
                            if(rgb === 'color-rgb(209,72,65)')
                                color = "attention";
                        }
                    }
               }

               if(element.data["text-align"]){
                   cAlign = element.data["text-align"];
               }

                data = {
                        "type": 'TextRun',
                        "text": eText,
                        "color": color,
                        "fontType": fontType,
                        "highlight" : highlight,
                        "isSubtle" : isSubtle,
                        "italic" : italic,
                        "size" : size,
                        "strikethrough": strikethrough,
                        "underline": underline,
                        "weight" : weight,
                        
                };
                jsonArr.push(data);
                setCardSummary(this.card, jsonArr, cAlign);
            }
          }
    }
}

const sendConfirmationTaskModuleWithTranslation = withTranslation()(SendConfirmationTaskModule);
export default sendConfirmationTaskModuleWithTranslation;
