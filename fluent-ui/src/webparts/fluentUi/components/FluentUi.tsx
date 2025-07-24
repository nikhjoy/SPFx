import * as React from 'react';
import { TextField, PrimaryButton, Dropdown, IDropdownOption, Checkbox, Toggle, DatePicker, DayOfWeek } from '@fluentui/react';

const dropdownOptions: IDropdownOption[] = [
  { key: 'one', text: 'Option One' },
  { key: 'two', text: 'Option Two' },
  { key: 'three', text: 'Option Three' },
];

const FluentUiDemo: React.FC = () => {
  const [textValue, setTextValue] = React.useState('');
  const [dropdownValue, setDropdownValue] = React.useState<string | undefined>();
  const [isChecked, setIsChecked] = React.useState(false);
  const [toggleValue, setToggleValue] = React.useState(false);
  const [dateValue, setDateValue] = React.useState<Date | undefined>(undefined);


  const handleButtonClick = () => {
    alert(`Text: ${textValue}, Dropdown: ${dropdownValue}, Checked: ${isChecked}, Toggle: ${toggleValue}, Date: ${dateValue}`);
  };

  return (
    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { width: 400, padding: 20 } }}>
      <h2>Fluent UI Components Demo</h2>

      <TextField
        label="Enter text"
        value={textValue}
        onChange={(_, newValue) => setTextValue(newValue || '')}
      />

      <Dropdown
        label="Select an option"
        options={dropdownOptions}
        selectedKey={dropdownValue}
        onChange={(_, option) => setDropdownValue(option?.key as string)}
      />

      <Checkbox
        label="I agree"
        checked={isChecked}
        onChange={(_, checked) => setIsChecked(!!checked)}
      />

      <Toggle
        label="Toggle me"
        checked={toggleValue}
        onChange={(_, checked) => setToggleValue(!!checked)}
      />

<DatePicker
  label="Pick a date"
  firstDayOfWeek={DayOfWeek.Sunday}
  value={dateValue}
  onSelectDate={(date) => setDateValue(date ?? undefined)}
/>

      <PrimaryButton text="Submit" onClick={handleButtonClick} />
    </Stack>
  );
};

export default FluentUiDemo;
