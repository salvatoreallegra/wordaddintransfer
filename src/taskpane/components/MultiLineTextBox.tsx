import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
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

export const MultiLineTextBox: React.FunctionComponent = () => {
  const [multiline, { toggle: toggleMultiline }] = useBoolean(false);
  const onChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    const newMultiline = newText.length > 50;
    if (newMultiline !== multiline) {
      toggleMultiline();
    }
  };
  return <TextField label="Standard" multiline rows={3} onChange={onChange} />;
};
