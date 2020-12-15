import React, { useState } from 'react'
import './App.css';

import { useMsal, useMsalAuthentication, AuthenticatedTemplate } from "@azure/msal-react";
import { Spinner, Label, ProgressIndicator, SelectionMode, GroupHeader, DetailsList, Fabric, MessageBar, MessageBarType, PrimaryButton, Stack, DefaultButton, Separator, Dropdown, Slider, Panel, PanelType, } from '@fluentui/react'
import { Icon } from '@fluentui/react/lib/Icon';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';

import { TagPicker } from 'office-ui-fabric-react/lib/Pickers';

initializeIcons();


function TimeItem({ dismissPanel, _add, projects, item }) {

  const [error, setError] = useState(null)
  console.log(item)

  const [input, handleInputChange] = useState({
    'hours': item ? item.hours : 1,
    'status': item ? item.status : "complete",
    'day': item ? item.day : 0,
    'project': item ? item.project : ""
  })

  function _onChange(e, val) {
    handleInputChange({
      ...input,
      [e.target.name]: val
    })
  }

  function add() {
    _add(input)
    //setError(null)
    //_fetchit('/api/store/inventory', 'POST', {}, result._id ? { _id: result._id, ...input } : input).then(succ => {
    //  console.log(`created success : ${JSON.stringify(succ)}`)
    //navTo("/MyBusiness")
    //dismissPanel()
    //}, err => {
    //  console.error(`created failed : ${err}`)
    //  setError(`created failed : ${err}`)
    //})
  }
  /*
    const testTags = [
      'Alpha Project',
      'Beta Project',
      'Gamma Project',
      'Delta Project',
      'Epsilon Project',
      'Zeta Project',
      'Eta Project',
      'Theta Project',
      'Iota Project',
      'Kappa Project',
      'Lambda Project',
      'Mu Project',
      'Nu Project',
      'Xi Project'
    ].map(item => ({ key: item, name: item }));
  */
  const testTags = projects.map(item => ({ key: item, name: item }));
  console.log(testTags)

  function onItemSelected(i) {
    _onChange({ target: { name: "project" } }, i.name)
    return i
  }
  return (
    <Stack tokens={{ childrenGap: 15 }} >

      <Dropdown label="Day" defaultSelectedKey={input.day} onChange={(e, i) => _onChange({ target: { name: "day" } }, i.key)} options={[{ key: 0, text: "Sunday" }, { key: 1, text: "Monday" }, { key: 2, text: "Tuesday" }, { key: 3, text: "Wednesday" }, { key: 4, text: "Thursday" }, { key: 5, text: "Friday" }, { key: 6, text: "Saturday" }]} />

      <Label >Project Search</Label>
      <TagPicker
        label="Project"
        inputProps={{ defaultVisibleValue: input.project }}
        removeButtonAriaLabel="Remove"
        onResolveSuggestions={(filterText, tagList) => {
          return filterText
            ? testTags.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) >= 0)
            : []
        }}
        getTextFromItem={(item) => item.name}
        pickerSuggestionsProps={{
          suggestionsHeaderText: 'Suggested projects',
          noResultsFoundText: 'No projects found',
        }}
        itemLimit={2}
        onItemSelected={onItemSelected}
      />

      <Slider
        label="Hours"
        min={1}
        max={10}
        step={1}
        defaultValue={input.hours}
        showValue={true}
        onChange={(val) => _onChange({ target: { name: "hours" } }, val)}
        snapToStep
      />

      <Dropdown label="Assosiated with outlook" defaultSelectedKey={input.status} onChange={(e, i) => _onChange({ target: { name: "status" } }, i.key)} options={[{ key: "title", text: "Title includes" }, { key: "category", text: "Catorgorised" }]} />

      {error &&
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false} truncated={true}>
          {error}
        </MessageBar>
      }
      <Stack horizontal tokens={{ childrenGap: 5 }}>
        <PrimaryButton text="Save" onClick={add} allowDisabledFocus disabled={false} />
        <DefaultButton text="Cancel" onClick={dismissPanel} allowDisabledFocus disabled={false} />
      </Stack>

    </Stack>
  )
}

const iconClass = mergeStyles({
  fontSize: 20,
  height: 20,
  width: 20,
  marginRight: '5px',
});

const flexrow = mergeStyles({
  width: "100%",
  display: "flex",
  flexDirection: "row",
  justifyContent: "left"
})

function App() {

  const [panel, setPanel] = React.useState({ open: false })
  const [importcal, setImportcal] = React.useState(0)
  const [projects, setProjects] = React.useState([
    'Alpha Project',
    'Beta Project',
    'Gamma Project',
    'Delta Project',
    'Epsilon Project',
    'Zeta Project',
    'Eta Project',
    'Theta Project',
    'Iota Project',
    'Kappa Project',
    'Lambda Project',
    'Mu Project',
    'Nu Project',
    'Xi Project'])
  const [entries, setEntries] = React.useState({
    hours: 0,
    groups: [
      { key: 'sun', name: 'Sunday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'mon', name: 'Monday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'tue', name: 'Tuesday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'wed', name: 'Wednesday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'thu', name: 'Thursday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'fri', name: 'Friday', startIndex: 0, count: 0, level: 0, hours: 0 },
      { key: 'sat', name: 'Saturday', startIndex: 0, count: 0, level: 0, hours: 0 }
    ],
    items: []
  })

  const { login, result, error } = useMsalAuthentication("redirect");
  const { instance, accounts, inProgress } = useMsal();

  function openTimeItem(editid) {
    setPanel({ open: true, editid })
  }

  function dismissPanel() {
    setPanel({ open: false })
  }

  function _sumbit() {

  }

  function add_in_order(new_val, array) {
    for (let i = 0; i < array.length; i++) {
      if (new_val.day <= array[i].day) {
        return [...array.slice(0, i), new_val, ...array.slice(i)]
      }
    }
    return [...array, new_val]
  }

  function add(item) {
    additems([item])
    dismissPanel()
  }

  function additems(newitems) {

    const ps = new Set(projects)

    setEntries((prevState) => {

      let items = [...prevState.items]
      let newhours = 0
      for (const n of newitems) {
        ps.add(n.project)
        newhours += n.hours
        items = add_in_order(n, items)
      }

      //console.log(items)

      let groups = newitems.reduce((g, i) => {
        g[i.day].count++; g[i.day].hours += i.hours
        for (let a = i.day + 1; a < g.length; a++) {
          g[a].startIndex++
        }
        //console.log(g)
        return g


      }, [...prevState.groups])

      console.log(items)
      console.log(groups)
      return { items, groups, hours: prevState.hours + newhours }

    })
    setProjects(Array.from(ps))
  }

  function _callAPI() {
    setImportcal(1)
    if (accounts && accounts.length > 0) {
      instance.acquireTokenSilent({
        scopes: [process.env.REACT_APP_SERVER_SCOPE],
        account: accounts[0]
      }).then((response) => {
        if (response) {

          fetch(process.env.REACT_APP_SERVER_CALURL, {
            headers: {
              'Authorization': `Bearer ${response.accessToken}`,
              'Content-Type': 'application/json'
            },
            mode: 'cors'
          }).then(response => response.json())
            .then(data => {

              additems(data.events.map(i => {
                const d = new Date(i.startTime.substr(6, 4), i.startTime.substr(3, 2) - 1, i.startTime.substr(0, 2), 1)
                console.log(`adding ${i.subject}  -- ${i.startTime} -- ${d.getDay()} -- ${i.durationInMinutes}`)
                return {
                  project: i.subject,
                  hours: Number.parseFloat((i.durationInMinutes / 60).toFixed(1)),
                  day: d.getDay(),
                  categories: i.categories
                }
              })//.filter(f => f.day >= 0 && f.day <= 4)
              )
              setImportcal(0)
            })
        }
      })
    }
  }

  return (
    <Fabric>
      <AuthenticatedTemplate>
        <main id="mainContent" data-grid="container" >

          <nav className="header">

            <div className="logo" style={{ padding: "6px 0" }}>
              <Icon iconName="TimeEntry" style={{ fontSize: 23, margin: '0 15px', color: 'deepskyblue' }} />
            </div>
            <div className="logo" style={{ padding: "8px 0" }}>
              <div style={{ fontSize: 15 }}>Time Recording Assistant</div>
            </div>
            <input className="menu-btn" type="checkbox" id="menu-btn" />
            <label className="menu-icon" htmlFor="menu-btn"><span className="navicon"></span></label>
            <ul className="menu">
              <li><a >Time Entry</a></li>
              <li><a >My Projects</a></li>
              <li><a >My Analytics</a></li>
              <li><a onClick={() => instance.logout()}>Logout</a></li>
            </ul>
          </nav>

          <div style={{ "height": "43px", "width": "100%" }} />


          <Stack className="wrapper" tokens={{ childrenGap: 10, padding: 10, maxWidth: "900px" }}>

            <Panel
              headerText="New Time entry"
              isOpen={panel.open}
              onDismiss={dismissPanel}
              type={PanelType.small}
              // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
              closeButtonAriaLabel="Close"
            >
              {panel.open &&
                <TimeItem dismissPanel={dismissPanel} _add={add} item={panel.item} projects={projects} />
              }
            </Panel>

            <Stack horizontal tokens={{ childrenGap: 40, padding: 10 }}>
              <div style={{ height: "65px", width: "65px" }} className="ms-BrandIcon--icon96 ms-BrandIcon--outlook"></div>

              <DefaultButton style={{ margin: '15px 15px' }} iconProps={{ iconName: 'CalendarWeek' }} text="Import from events" onClick={_callAPI} disabled={(importcal !== 0)} />
              {importcal === 1 &&
                <Spinner style={{ marginLeft: "0px" }} label="loading..." />
              }
            </Stack>

            <Separator></Separator>
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Create Time Entry" styles={{ root: { width: 180 } }} onClick={() => openTimeItem()} />

            <DetailsList

              items={entries.items}
              groups={entries.groups}
              columns={[
                { key: 'day', name: `Week (${entries.hours} hrs)`, minWidth: 100, maxWidth: 200, isResizable: false },
                { key: 'project', name: 'Project', fieldName: 'project', minWidth: 100, maxWidth: 200, onRender: (i) => <Label>{i.project}</Label> },
                { key: 'hours', name: 'Hours', fieldName: 'hours', minWidth: 100, maxWidth: 200, onRender: (i) => <Label>{i.hours} hrs</Label> },
                {
                  key: 'cat', name: 'Outlook Categories', fieldName: 'categories', minWidth: 100, maxWidth: 200, onRender: (i) =>
                    <div>{i.categories && i.categories.map((c, i) =>
                      <div key={i} className={flexrow}>
                        <Icon iconName="Tag" className={iconClass} style={{ color: "red" }} />
                        <div>{c}</div>
                      </div>)}
                    </div>
                }
              ]}
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              ariaLabelForSelectionColumn="Toggle selection"
              checkButtonAriaLabel="Row checkbox"
              selectionMode={SelectionMode.none}
              groupProps={{
                showEmptyGroups: true,
                onRenderHeader: (item) => <GroupHeader onRenderGroupHeaderCheckbox={false} {...item} onRenderTitle={(i) =>
                  <div className='ms-GroupHeader-title' role="gridcell">
                    <span>{i.group.name}  ({i.group.hours} hrs
                    {i.group.hasMoreData && '+'})
                </span>
                  </div>} />
              }}
              onActiveItemChanged={(i) => setPanel({ open: true, item: i })}
              compact={false}
            />

            <ProgressIndicator label={`${Number.parseFloat((entries.hours / 40) * 100).toFixed(0)}% Complete`} percentComplete={entries.hours / 40} barHeight={entries.hours} />

            <Stack.Item align="end">
              <DefaultButton text="Submit" onClick={_sumbit} allowDisabledFocus disabled={entries.hours < 40} />
            </Stack.Item>
          </Stack>

        </main>
      </AuthenticatedTemplate>
    </Fabric>
  )
}

export default App;
