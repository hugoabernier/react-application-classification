import * as React from "react";
import {
  IClassificationHeaderProps,
  IClassificationHeaderState,
} from "./ClassificationHeader.types";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Link } from "office-ui-fabric-react/lib/Link";
import { Web } from "sp-pnp-js/lib/pnp";
import * as strings from "ClassificationExtensionApplicationCustomizerStrings";


export default class ClassificationHeader extends React.Component<IClassificationHeaderProps, IClassificationHeaderState> {
  constructor(props: IClassificationHeaderProps) {
    super(props);
    this.state = {
      isLoading: true,
      businessImpact: null
    };
  }

  public componentDidMount(): void {
    this.setState({
      isLoading: true,
      businessImpact: null
    });

    console.log("Retrieving property bags");
    const web: Web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.select("Title", "AllProperties")
      .expand("AllProperties")
      .get()
      .then(r => {
        console.log("Property bags",r);
        var businessImpact: string = this.props.DefaultClassification;
        console.log("DefaultClassification",businessImpact);

        // handle the default situation where there is no classification
        if (r.AllProperties && r.AllProperties[this.props.ClassificationPropertyBag]) {
          businessImpact = r.AllProperties[this.props.ClassificationPropertyBag];
        }

        console.log("businessImpact",businessImpact);

        this.setState({
          isLoading: false,
          businessImpact: businessImpact
        });
        console.log("All properties results", r);
      });
  }

  public render(): React.ReactElement<IClassificationHeaderProps> {
    console.log("Render");
    // get the business impact from the state
    let { businessImpact } = this.state;

    // ge the default handling URL
    let handlingUrl: string = this.props.DefaultHandlingUrl;

    // change this switch statement to suit your security classification
    var barType: MessageBarType;
    switch (businessImpact) {
      case "MBI":
        // if you'd like to display a different URL per classification, override the handlingUrl variable here
        // handlingUrl = "/SitePages/Handling-instructions-MBI.aspx"
        barType = MessageBarType.warning;
        break;
      case "HBI":
        barType = MessageBarType.severeWarning;
        break;
      case "LBI":
        barType = MessageBarType.info;
        break;
      default:
        barType = undefined;
    }

    // if no security classification, do not display a header
    if (barType === undefined) {
      console.log("Bar type is undefined");
      return null;
    }

    return (
      <MessageBar
        messageBarType={barType}
      >
        {strings.ClassifactionMessage.replace("{0}",this.state.businessImpact)}
        {handlingUrl && handlingUrl !== undefined && handlingUrl !== "" ?
          <Link
            href={handlingUrl}

          > {strings.HandlingMessage}</Link>
          : null
        }
      </MessageBar>
    );
  }
}
