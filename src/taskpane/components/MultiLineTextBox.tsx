import * as React from "react";
import { useState } from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Button, ButtonType } from "office-ui-fabric-react";

import { useBoolean } from "@uifabric/react-hooks";
import { lorem } from "@uifabric/example-data";
import { Stack, IStackProps, IStackStyles } from "office-ui-fabric-react/lib/Stack";

const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const stackTokens = { childrenGap: 50 };
const dummyText: string = lorem(100);
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } }
};

export const MultiLineTextBox: React.FunctionComponent = props => {
  const [multiline, { toggle: toggleMultiline }] = useBoolean(false);
  //const [count, setCount] = useState(0);
  // const onChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
  //   const newMultiline = newText.length > 50;
  //   if (newMultiline !== multiline) {
  //     toggleMultiline();
  //   }
  // };

  //function add() {}
  return (
    <div>
      <TextField label="Enter FetchXML" multiline rows={3} /* onChange={onChange}*/ />
      {/* <Button
        className="ms-welcome__action"
        buttonType={ButtonType.hero}
        iconProps={{ iconName: "ChevronRight" }}
        onClick={add}
      >
        Add
      </Button> */}
    </div>
  );
};
