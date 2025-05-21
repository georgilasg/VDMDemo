import * as React from 'react';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import { Pivot, PivotItem, Stack, TextField, Text, PrimaryButton, DefaultButton } from '@fluentui/react';
import { IChoiceGroupOption, ChoiceGroup, ComboBox, IComboBoxOption, } from '@fluentui/react';
import { Dropdown, IDropdownOption, DatePicker } from '@fluentui/react';

export interface IVdmDemoFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: (wasSaved: boolean) => void;
}

const VdmDemoForm: React.FC<IVdmDemoFormProps> = (props) => {
  // State variables for form fields
  const [title, setTitle] = React.useState<string>('');
  const [description, setDescription] = React.useState<string>('');
  const [status, setStatus] = React.useState<string>('Not Started');
  const [priority, setPriority] = React.useState<string>('Medium');
  const [dueDate, setDueDate] = React.useState<Date | null>(null);
  const [assignedTo, setAssignedTo] = React.useState<string | undefined>(undefined);
  const [category, setCategory] = React.useState<string>('Development');
  const [completion, setCompletion] = React.useState<number>(0);
  const [comments, setComments] = React.useState<string>('');
  const [users, setUsers] = React.useState<IComboBoxOption[]>([]);
  const [tags, setTags] = React.useState<string[]>([]);
  const [activeTab, setActiveTab] = React.useState<string>('general');
  const [documentTypeOptions, setDocumentTypeOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedDocumentType, setSelectedDocumentType] = React.useState<string | undefined>(undefined);
  const [selectedDocumentSubTypes, setSelectedDocumentSubTypes] = React.useState<string[]>([]);
  const [filteredDocumentSubTypeOptions, setFilteredDocumentSubTypeOptions] = React.useState<IDropdownOption[]>([]);

  const listItem = "VDMDemo";
  const listDocumentType = "DocumentType";
  const listDocumentSubType = "DocumentSubType";
  const listItemEntityType = "SP.Data.VDMDemoListItem";

  // Dropdown options for status, priority, and category
  const statusOptions: IDropdownOption[] = [
    { key: 'Not Started', text: 'Not Started' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Completed', text: 'Completed' },
    { key: 'On Hold', text: 'On Hold' },
  ];

  const priorityOptions: IChoiceGroupOption[] = [
    { key: 'Low', text: 'Low' },
    { key: 'Medium', text: 'Medium' },
    { key: 'High', text: 'High' },
  ];

  const categoryOptions: IDropdownOption[] = [
    { key: 'Development', text: 'Development' },
    { key: 'Testing', text: 'Testing' },
    { key: 'Documentation', text: 'Documentation' },
    { key: 'Support', text: 'Support' },
  ];

  //For the Tags Control, Options and CSS
  const tagOptions: string[] = ['Frontend', 'Backend', 'Database', 'Testing'];
  const checkboxStyle = {
    marginRight: '8px',
    transform: 'scale(1.5)',  // Increase the size of the checkbox
    cursor: 'pointer'
  };

  const labelStyle = {
    display: 'inline-flex',
    alignItems: 'center',
    padding: '5px 12px',
    margin: '5px',
    border: '2px solid #0078d4',
    borderRadius: '8px',
    backgroundColor: '#f3f2f1',
    cursor: 'pointer',
    transition: 'all 0.2s ease',
  };

  const activeLabelStyle = {
    ...labelStyle,
    backgroundColor: '#0078d4',
    color: '#fff',
    fontWeight: 'bold',
  };

  interface IDocumentType {
    Id: number;
    Title: string;
  }

  interface IDocumentSubType {
    Id: number;
    Title: string;
    DocTypeId: number;
  }

  interface IUser {
    Id: number;
    Title: string;
    PrincipalType: number;
  }

  // Fetch data From the list item
  const fetchItem = async (): Promise<void> => {
    const response = await props.context.spHttpClient.get(
      `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listItem}')/items(${props.context.itemId})`,
      SPHttpClient.configurations.v1
    );
    if (response.ok) {
      const item = await response.json();

      setTitle(item.Title || '');
      setDescription(item.Description || '');
      setStatus(item.Status || 'Not Started');
      setPriority(item.Priority || 'Medium');
      setDueDate(item.DueDate ? new Date(item.DueDate) : null);
      setAssignedTo(item.AssignedToId?.toString() || '');
      setCategory(item.Category || '');
      setCompletion(item.CompletionPercentage || 0);
      setComments(item.Comments || '');
      setTags(item.Tags?.results || []);
      setSelectedDocumentType(item.DocumentTypeId?.toString() || '');
      setSelectedDocumentSubTypes(item.DocumentSubTypesId?.results?.map((id: number) => id.toString()) || []);
    }
  };

  // Fetch data for DocumentTypes (single lookup)
  const fetchDocumentTypes = async (): Promise<void> => {
    try {
      const response = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listDocumentType}')/items`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        const options = data.value.map((item: IDocumentType) => ({
          key: item.Id.toString(),
          text: item.Title,
        }));
        setDocumentTypeOptions(options);
        console.log('DocumentTypes:', options);
      }
    } catch (error) {
      console.error('Error fetching DocumentTypes:', error);
    }
  };

  // Use an inline async function to handle the fetch
  const fetchUsers = async (): Promise<void> => {
    try {
      const response = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/siteusers`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const userOptions = data.value
          .filter((user: IUser) => user.PrincipalType === 1)
          .map((user: IUser) => ({
            key: user.Id.toString(),
            text: user.Title,
          }));
        setUsers(userOptions);
        console.log('Fetched users:', userOptions);
      } else {
        console.error('Error fetching users:', response.statusText);
      }
    } catch (error) {
      console.error('Error fetching users:', error);
    }
  };

  // Fetch data for DocumentSubTypes (multi-lookup)
  const fetchDocumentSubTypes = async (selectedTypeId: string): Promise<void> => {
    try {
      const response = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listDocumentSubType}')/items`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        // Filter subtypes based on the selected document type
        const filteredOptions = data.value
          .filter((item: IDocumentSubType) => item.DocTypeId?.toString() === selectedTypeId)
          .map((item: IDocumentSubType) => ({
            key: item.Id.toString(),
            text: item.Title
          }));

        setFilteredDocumentSubTypeOptions(filteredOptions);
        console.log('Filtered DocumentSubTypes:', filteredOptions);
      }
    } catch (error) {
      console.error('Error fetching DocumentSubTypes:', error);
    }
  };

  React.useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        await fetchUsers();
        await fetchDocumentTypes();

        if (props.displayMode === FormDisplayMode.Edit && props.context.itemId) {
          await fetchItem();
        }
      } catch (err) {
        console.error('Error during useEffect fetches:', err);
      }
    };

    fetchData().catch((error) => console.error('Error in async function:', error));
  }, [props.context]);

  const handleDocumentTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setSelectedDocumentType(option.key as string);
      setSelectedDocumentSubTypes([]); // Clear the previous selection
      // Fetch the filtered DocumentSubTypes based on the selected type
      fetchDocumentSubTypes(option.key as string).catch((error) => console.error('Error in async function:', error));
    }
  };

  // Handle form submission
  const handleSubmit = async (): Promise<void> => {
    try {

      const webUrl = props.context.pageContext.web.absoluteUrl;
      let requestUrl: string;
      let method: 'POST' | 'PATCH';

      // Prepare the item payload with metadata
      const item = {
        __metadata: { type: listItemEntityType },
        Title: title,
        Description: description,
        Status: status,
        DueDate: dueDate ? dueDate.toISOString() : null,
        AssignedToId: assignedTo ? parseInt(assignedTo) : null,
        Category: category,
        CompletionPercentage: completion,
        Priority: priority,
        Tags: { results: tags },
        DocumentTypeId: selectedDocumentType ? parseInt(selectedDocumentType) : null,
        DocumentSubTypesId: { results: selectedDocumentSubTypes.map(Number) },
        Comments: comments
      };

      const headers = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        'IF-MATCH': '*'
      };

      if (props.displayMode === FormDisplayMode.Edit && props.context.itemId) {
        // Update existing item
        method = 'PATCH';
        requestUrl = `${webUrl}/_api/web/lists/getbytitle('${listItem}')/items(${props.context.itemId})`;
      }
      else {
        // Create a new item
        method = 'POST';
        requestUrl = `${webUrl}/_api/web/lists/getbytitle('${listItem}')/items`;
      }

      // Make the POST request to create a new item
      const response = await props.context.spHttpClient.fetch(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          method: method,
          headers: headers,
          body: JSON.stringify(item)
        }
      );

      if (response.ok) {
        alert(props.displayMode === FormDisplayMode.Edit ? 'Item updated successfully!' : 'Item created successfully!');

        // Clear the form fields after successful submission in case we wanted to keep the form up
        /* setTitle('');
        setDescription('');
        setStatus('Not Started');
        setPriority('Medium');
        setDueDate(null);
        setAssignedTo(undefined);
        setCategory('Development');
        setCompletion(0);
        setComments('');
        setTags([]);
        setActiveTab('general'); */

        //Close the form
        props.onClose(true);
      } else {
        const errorData = await response.json();
        const errorMessage = errorData?.error?.message?.value || JSON.stringify(errorData);
        alert(`Error saving item: ${errorMessage}`);
        console.error('Error response:', errorData);
      }
    } catch (error) {
      alert(`An unexpected error occurred: ${error.message}`);
      console.error('Unexpected error:', error);
    }
  };


  return (
    <div style={{ padding: '20px' }}>
      <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey || 'general')}>
        <PivotItem headerText="General Info" itemKey="general">
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Title:</Text>
            <Stack grow> <TextField
              value={title}
              onChange={(e, v) => setTitle(v || '')}
              placeholder="Enter the title"
              required
            />
            </Stack>
          </Stack>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Description:</Text>
            <Stack grow> <TextField multiline rows={3} value={description} onChange={(e, v) => setDescription(v || '')} />
            </Stack>
          </Stack>
        </PivotItem>



        <PivotItem headerText="Details" itemKey="details">
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Priority:</Text>
            <Stack grow>
              <ChoiceGroup
                selectedKey={priority}
                options={priorityOptions}
                onChange={(event, option) => setPriority(option?.key || 'Medium')}
                styles={{ flexContainer: { display: 'flex', flexDirection: 'row', gap: '10px' } }}
              />
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }} style={{ width: '50%' }}>
              <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px', marginRight: '5px' }}>Status:</Text>
              <Stack grow>
                <Dropdown
                  selectedKey={status}
                  options={statusOptions}
                  onChange={(e, option) => setStatus(option?.key as string)}
                />
              </Stack>
            </Stack>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }} style={{ width: '50%' }}>
              <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px', marginRight: '5px' }}>Category:</Text>
              <Stack grow>
                <Dropdown
                  selectedKey={category}
                  options={categoryOptions}
                  onChange={(e, option) => setCategory(option?.key as string)}
                />
              </Stack>
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }} style={{ width: '50%' }}>
              <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px', marginRight: '5px' }}>Due Date:</Text>
              <Stack grow>
                <DatePicker
                  value={dueDate || undefined}
                  onSelectDate={(date) => setDueDate(date || null)}
                />
              </Stack>
            </Stack>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }} style={{ width: '50%' }}>
              <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px', marginRight: '5px' }}>Completion (%):</Text>
              <Stack grow>
                <div style={{ display: 'flex', gap: '20px' }}>
                  <TextField
                    type="number"
                    min={0}
                    max={100}
                    value={completion.toString()}
                    onChange={(e, v) => setCompletion(Number(v) || 0)}
                    style={{ width: '150px' }}
                  />
                </div>
              </Stack>
            </Stack>
          </Stack>
        </PivotItem>

        <PivotItem headerText="Assignment" itemKey="assignment">

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Assigned To:</Text>
            <Stack grow>
              <ComboBox
                selectedKey={assignedTo}
                placeholder="Select a user"
                allowFreeform
                autoComplete="on"
                options={users}
                onChange={(e, option) => setAssignedTo(option?.key?.toString())}
              />
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Tags:</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
              {tagOptions.map(tag => (
                <label
                  key={tag}
                  style={tags.includes(tag) ? activeLabelStyle : labelStyle}
                  onClick={() => {
                    if (tags.includes(tag)) {
                      setTags(tags.filter(t => t !== tag));
                    } else {
                      setTags([...tags, tag]);
                    }
                  }}
                >
                  <input
                    type="checkbox"
                    value={tag}
                    checked={tags.includes(tag)}
                    onChange={(e) => {
                      if (e.target.checked) {
                        setTags([...tags, tag]);
                      } else {
                        setTags(tags.filter(t => t !== tag));
                      }
                    }}
                    style={checkboxStyle}
                  />
                  {tag}
                </label>
              ))}
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Document Type:</Text>
            <Stack grow>
              <Dropdown
                placeholder="Select Document Type"
                selectedKey={selectedDocumentType}
                options={documentTypeOptions}
                onChange={handleDocumentTypeChange}
              />
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Document SubTypes:</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
              {filteredDocumentSubTypeOptions.map(option => (
                <label
                  key={option.key}
                  style={selectedDocumentSubTypes.includes(option.key as string) ? activeLabelStyle : labelStyle}
                  onClick={() => {
                    const id = option.key as string;
                    if (selectedDocumentSubTypes.includes(id)) {
                      setSelectedDocumentSubTypes(selectedDocumentSubTypes.filter(item => item !== id));
                    } else {
                      setSelectedDocumentSubTypes([...selectedDocumentSubTypes, id]);
                    }
                  }}
                >
                  <input
                    type="checkbox"
                    value={option.key}
                    checked={selectedDocumentSubTypes.includes(option.key as string)}
                    onChange={(e) => {
                      const id = option.key as string;
                      if (e.target.checked) {
                        setSelectedDocumentSubTypes([...selectedDocumentSubTypes, id]);
                      } else {
                        setSelectedDocumentSubTypes(selectedDocumentSubTypes.filter(item => item !== id));
                      }
                    }}
                    style={checkboxStyle}
                  />
                  {option.text}
                </label>
              ))}
            </Stack>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '10px', width: '100%' }}>
            <Text variant="mediumPlus" style={{ width: '150px', minWidth: '150px' }}>Comments:</Text>
            <Stack grow>
              <TextField multiline rows={3} value={comments} onChange={(e, v) => setComments(v || '')} />
            </Stack>
          </Stack>
        </PivotItem>
      </Pivot>
      <Stack horizontal horizontalAlign="start" tokens={{ childrenGap: 10 }} style={{ marginTop: '20px' }}>
        <PrimaryButton text="Submit" onClick={handleSubmit} />
        <DefaultButton text="Cancel" onClick={() => props.onClose(false)} />
      </Stack>
    </div>
  );
};

export default VdmDemoForm;

/* 
const LOG_SOURCE: string = 'VdmDemoForm';
export default class VdmDemoForm extends React.Component<IVdmDemoFormProps> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: VdmDemoForm mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: VdmDemoForm unmounted');
  }

  public render(): React.ReactElement<IVdmDemoFormProps> {
    return <div className={styles.vdmDemoForm} />;
  }
} */
