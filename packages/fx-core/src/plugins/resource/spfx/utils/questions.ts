import { Question } from "@microsoft/teamsfx-api";
import { yeomanScaffoldEnabled } from "../../../../core/globalVars";

export enum SPFXQuestionNames {
  framework_type = "spfx-framework-type",
  webpart_name = "spfx-webpart-name",
  webpart_desp = "spfx-webpart-desp",
  component_type = "spfx-component-type",
  ace_type = "spfx-ace-type",
}

export const frameworkQuestion: Question = {
  type: "singleSelect",
  name: SPFXQuestionNames.framework_type,
  title: "Framework",
  staticOptions: yeomanScaffoldEnabled()
    ? [
        { id: "react", label: "React" },
        { id: "none", label: "None" },
        { id: "minimal", label: "Minimal" },
      ]
    : [
        { id: "react", label: "React" },
        { id: "none", label: "None" },
      ],
  placeholder: "Select an option",
  default: "react",
};

export const componentTypeQuestion: Question = {
  type: "singleSelect",
  name: SPFXQuestionNames.component_type,
  title: "Component Type",
  staticOptions: [
    { id: "webpart", label: "WebPart" },
    { id: "adaptiveCardExtension", label: "Adaptive Card Extension" },
  ],
  placeholder: "Select an option",
  default: "webpart",
};

export const aceTypeQuestion: Question = {
  type: "singleSelect",
  name: SPFXQuestionNames.ace_type,
  title: "Adaptive Card Extension Type",
  staticOptions: [
    { id: "Basic", label: "Basic Card Template" },
    { id: "PrimaryText", label: "Primary Text Template" },
    { id: "Image", label: "Image Card Template" },
  ],
  placeholder: "Select an option",
  default: "webpart",
};

export const webpartNameQuestion: Question = {
  type: "text",
  name: SPFXQuestionNames.webpart_name,
  title: "Solution Name",
  default: "helloworld",
  validation: {
    pattern: "^[a-zA-Z_][a-zA-Z0-9_]*$",
  },
};

export const webpartDescriptionQuestion: Question = {
  type: "text",
  name: SPFXQuestionNames.webpart_desp,
  title: "Solution Description",
  default: "helloworld description",
  validation: {
    required: true,
  },
};
