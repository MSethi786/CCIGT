// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";

var gg = 
	 [
		{
			"type": "TextRun",
			"text": "**Hello**",
			"wrap": true
		}
	];

var dd = [
    {
      "type": "TextRun",
      "text": "We support colors,",
      "color": "good"
    },
    {
      "type": "TextRun",
      "text": " both regular and subtle. ",
      "isSubtle": true
    },
    {
      "type": "TextRun",
      "text": "Text ",
      "size": "small"
    },
    {
      "type": "TextRun",
      "text": "of ",
      "size": "medium"
    },
    {
      "type": "TextRun",
      "text": "all ",
      "size": "large"
    },
    {
      "type": "TextRun",
      "text": "sizes! ",
      "size": "extraLarge"
    },
    {
      "type": "TextRun",
      "text": "Light weight text. ",
      "weight": "lighter"
    },
    {
      "type": "TextRun",
      "text": "Bold weight text. ",
      "weight": "bolder"
    },
    {
      "type": "TextRun",
      "text": "Highlights. ",
      "highlight": true
    },
    {
      "type": "TextRun",
      "text": "Italics. ",
      "italic": true
    },
    {
      "type": "TextRun",
      "text": "Strikethrough. ",
      "strikethrough": true
    },
    {
      "type": "TextRun",
      "text": "Monospace too!",
      "fontType": "monospace"
    }
  ];

export const getInitAdaptiveCard = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": titleTextAsString,
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": "",
                    "size": "Stretch",
                    "width": "400px",
                    "altText": ""
                },
                {
                    "type": "RichTextBlock",
                    "inlines": ""
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.3"
        }
    );
}

export const getCardTitle = (card: any) => {
    return card.body[0].text;
}

export const setCardTitle = (card: any, title: string) => {
    card.body[0].text = title;
}

export const getCardImageLink = (card: any) => {
    return card.body[1].url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    card.body[1].url = imageLink;
}

// export const getCardSummary = (card: any) => {
//     return card.body[2].text;
// }

// export const setCardSummary = (card: any, summary?: string) => {
//     card.body[2].text = summary;
//     console.log(summary);
// }

// export const getCardSummary1 = (card: any) => {
//     return card.body[3].text;
// }

// export const setCardSummary1 = (card: any, summary?: string) => {
//     card.body[3].text = summary;
// }

export const getCardSummary = (card: any) => {
    return card.body[2].text;
}

export const setCardSummary = (card: any, richText?: any, align?: any) => {
   // console.log("richText" + richText);
    card.body[2].horizontalAlignment = align;
    card.body[2].inlines = richText;
}

export const getCardAuthor = (card: any) => {
    return card.body[3].text;
}

export const setCardAuthor = (card: any, author?: string) => {
    card.body[3].text = author;
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
}

export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string) => {
    if (buttonTitle && buttonLink) {
        card.actions = [
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle,
                "url": buttonLink
            }
        ];
    } else {
        delete card.actions;
    }
}
