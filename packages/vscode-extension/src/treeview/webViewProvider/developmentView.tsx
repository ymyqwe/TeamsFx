import * as React from "react";
import "./tree.scss";
import addCapability_dark from "../../../media/dark/addCapability.svg";
import addCapability_light from "../../../media/light/addCapability.svg";
import addResources_dark from "../../../media/dark/addResources.svg";
import addResources_light from "../../../media/light/addResources.svg";
import { getCurrentTheme, Theme } from "./theme";
import { TreeItem } from "./treeItem";
import { localize } from "../../utils/localizeUtils";

export class DevelopmentView extends React.Component<any, any> {
  constructor(props: any) {
    super(props);
    this.state = {
      colorTheme: Theme.Dark,
      locked: false,
    };
  }

  componentDidMount() {
    window.addEventListener("message", this.receiveMessage, false);
    document.addEventListener("contextmenu", (e) => {
      e.preventDefault();
    });

    const targetNode = document.getElementsByTagName("body")[0];
    // Options for the observer (which mutations to observe)
    const config = { attributes: true };
    // Callback function to execute when mutations are observed
    const callback = (mutationsList: any) => {
      // Use traditional 'for loops' for IE 11
      for (const mutation of mutationsList) {
        if (mutation.type === "attributes") {
          const theme = getCurrentTheme(mutation);
          this.setState({
            colorTheme: theme,
          });
        }
      }
    };

    // Create an observer instance linked to the callback function
    const observer = new MutationObserver(callback);
    // Start observing the target node for configured mutations
    observer.observe(targetNode, config);
  }

  render() {
    return (
      <div>
        <TreeItem
          label={localize("teamstoolkit.commandsTreeViewProvider.createProjectTitleNew")}
          tooltip={localize("teamstoolkit.commandsTreeViewProvider.createProjectDescription")}
          icon="codicon codicon-new-folder"
          customized={false}
          disable={false}
          command="fx-extension.create"
        ></TreeItem>
        <TreeItem
          label={localize("teamstoolkit.commandsTreeViewProvider.samplesTitleNew")}
          tooltip={localize("teamstoolkit.commandsTreeViewProvider.samplesDescription")}
          icon="codicon codicon-library"
          customized={false}
          disable={false}
          command="fx-extension.openSamples"
        ></TreeItem>
        <TreeItem
          label={localize("teamstoolkit.commandsTreeViewProvider.addCapabilitiesTitleNew")}
          tooltip={localize("teamstoolkit.commandsTreeViewProvider.addCapabilitiesDescription")}
          icon={this.state.colorTheme === Theme.Dark ? addCapability_dark : addCapability_light}
          customized={true}
          disable={this.state.locked}
          command="fx-extension.addCapability"
        ></TreeItem>
        <TreeItem
          label={localize("teamstoolkit.commandsTreeViewProvider.addResourcesTitleNew")}
          tooltip={localize("teamstoolkit.commandsTreeViewProvider.addResourcesDescription")}
          icon={this.state.colorTheme === Theme.Dark ? addResources_dark : addResources_light}
          customized={true}
          disable={this.state.locked}
          command="fx-extension.update"
        ></TreeItem>
        <TreeItem
          label={localize("teamstoolkit.commandsTreeViewProvider.manifestEditorTitleNew")}
          tooltip={localize("teamstoolkit.commandsTreeViewProvider.manifestEditorDescription")}
          icon="codicon codicon-edit"
          customized={false}
          disable={false}
          command="fx-extension.openManifest"
        ></TreeItem>
      </div>
    );
  }

  receiveMessage = (event: any) => {
    const message = event.data.message;

    switch (message) {
      case "lockChanged":
        this.setState({ locked: event.data.data });
        break;
      default:
        break;
    }
  };
}
