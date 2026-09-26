/**
 * People / group field – prefers Graph/SPFx PeoplePicker when available;
 * falls back to text + email for environments without @pnp/spfx-controls.
 */

import * as React from 'react';
import { Stack, TextField } from '@fluentui/react';

export interface IPersonValue {
  displayName: string;
  email?: string;
  loginName?: string;
}

export interface IPeopleFieldProps {
  label: string;
  required?: boolean;
  value?: IPersonValue;
  onChange: (value: IPersonValue) => void;
  errorMessage?: string;
  /** When true, only name is required (email optional) */
  nameOnly?: boolean;
}

/**
 * Lightweight dual text fields. Replace with PeoplePicker control when
 * @pnp/spfx-controls-react is added to the project:
 *
 *   <PeoplePicker context={spfxContext} ... />
 */
export const PeopleField: React.FC<IPeopleFieldProps> = (props) => {
  const name = props.value?.displayName || '';
  const email = props.value?.email || '';

  return (
    <Stack horizontal tokens={{ childrenGap: 10 }}>
      <TextField
        label={props.label}
        required={props.required}
        value={name}
        errorMessage={props.errorMessage}
        onChange={(_, v) =>
          props.onChange({ displayName: v || '', email: props.value?.email, loginName: props.value?.loginName })
        }
        styles={{ root: { flex: 1 } }}
        placeholder="Display name"
      />
      {!props.nameOnly && (
        <TextField
          label={`${props.label} email`}
          type="email"
          value={email}
          onChange={(_, v) =>
            props.onChange({
              displayName: props.value?.displayName || '',
              email: v || '',
              loginName: props.value?.loginName
            })
          }
          styles={{ root: { flex: 1 } }}
          placeholder="user@contoso.com"
        />
      )}
    </Stack>
  );
};
