import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";

export const MultiLineTextBox: React.FunctionComponent = () => {
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
