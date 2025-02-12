import * as React from "react";
import { useState, useEffect } from "react";
import { GroupedList, IGroup } from "@fluentui/react/lib/GroupedList";
import { IColumn, DetailsRow, DetailsRowFields } from "@fluentui/react/lib/DetailsList";
import {
  Selection,
  SelectionMode,
  SelectionZone,
} from "@fluentui/react/lib/Selection";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { useConst } from "@fluentui/react-hooks";

export type Dataset = ComponentFramework.PropertyTypes.DataSet;
export interface IGroupedList {
  dataset: Dataset;
  context: ComponentFramework.Context<IInputs>;
}

interface GroupedRecord {
  [key: string]: {
    [key: string]: Array<{ [key: string]: string }>;
  };
}

export interface DynamicItem {
  [key: string]: string | number | boolean | undefined;
}

export function makeGroups(
  dataset: Dataset,
  level1: string,
  level2: string
): { groups: IGroup[]; items: DynamicItem[] } {
  const groups: IGroup[] = [];
  const items: DynamicItem[] = [];
  let startIndex = 0;

  // Create unique groups for level1
  const uniqueLevel1Values = new Set<string>();
  dataset.sortedRecordIds.forEach((id) => {
    const record = dataset.records[id];
    const value1 = record.getFormattedValue(level1);
    if (value1) {
      uniqueLevel1Values.add(value1);
    }
  });

  // Iterate over unique level1 values and create subgroups for level2
  uniqueLevel1Values.forEach((value1) => {
    const level2Groups: IGroup[] = [];
    const uniqueLevel2Values = new Set<string>();

    dataset.sortedRecordIds.forEach((id) => {
      const record = dataset.records[id];
      if (record.getFormattedValue(level1) === value1) {
        const value2 = record.getFormattedValue(level2);
        if (value2) {
          uniqueLevel2Values.add(value2);
        }
      }
    });

    // Generate Level 2 Groups
    uniqueLevel2Values.forEach((value2) => {
      const matchingItems = dataset.sortedRecordIds
        .map((id) => dataset.records[id])
        .filter(
          (record) =>
            record.getFormattedValue(level1) === value1 &&
            record.getFormattedValue(level2) === value2
        );

      const itemCount = matchingItems.length;

      matchingItems.forEach((record) => {
        const dynamicItem: DynamicItem = {
          ...Object.keys(record).reduce((acc, key) => {
            acc[key] = record.getFormattedValue(key) || undefined;
            return acc;
          }, {} as DynamicItem),
        };
        items.push(dynamicItem);
      });

      level2Groups.push({
        key: `${value1}-${value2}`,
        name: value2,
        startIndex: startIndex,
        count: itemCount,
        level: 1, // Nested level 2
        isCollapsed: true,
      });

      startIndex += itemCount;
    });

    // Create level1 groups, only adding level2 groups as children
    const itemCount = level2Groups.reduce((sum, group) => sum + group.count, 0);

    groups.push({
      key: value1,
      name: value1,
      startIndex: startIndex - itemCount,
      count: itemCount,
      level: 0, // Top level
      isCollapsed: true,
      children: level2Groups, // Attach level2 groups as children
    });
  });

  return { groups, items };
}

export function makeColumnAndItems(
  dataset: Dataset,
  level1: string,
  level2: string,
  groups: IGroup[]
): {
  items: DynamicItem[];
  columns: IColumn[];
} {
  
  const columns: IColumn[] = dataset.columns.filter((column) => column.name !== "ct_name").slice(3).map((column) => ({
    name: column.displayName,
    fieldName: column.name,
    minWidth: column.visualSizeFactor,
    key: column.name,
}));

  const items: DynamicItem[] = [];
  const groupedRecords: Record<string, Record<string, DynamicItem[]>> = {};

  dataset.sortedRecordIds.forEach((id) => {
    const record = dataset.records[id];
    const level1Value = record.getFormattedValue(level1) || "";
    const level2Value = record.getFormattedValue(level2) || "";

    if (!groupedRecords[level1Value]) {
      groupedRecords[level1Value] = {};
    }

    if (!groupedRecords[level1Value][level2Value]) {
      groupedRecords[level1Value][level2Value] = [];
    }

    groupedRecords[level1Value][level2Value].push({
      ...Object.fromEntries(
        dataset.columns.map((column) => [
          column.name,
          record.getFormattedValue(column.name) || "",
        ])
      ),
    });
  });

  Object.keys(groupedRecords).forEach((level1Key) => {
    const level1Group = groupedRecords[level1Key];

    Object.keys(level1Group).forEach((level2Key) => {
      level1Group[level2Key].forEach((item) => {
        items.push({
          ...item,
          [level1Key]: level1Key,
          [level2Key]: level2Key,
        });
      });
    });
  });

  return { items, columns };
}

export const GroupedListComp = ({ dataset, context }: IGroupedList & { context: ComponentFramework.Context<IInputs> }): JSX.Element => {
  const [groups, setGroups] = useState<IGroup[]>([]);
  const [items, setItems] = useState<DynamicItem[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [level1, setLevel1] = useState<string>(dataset.columns[0].name);
  const [level2, setLevel2] = useState<string>(dataset.columns[1].name);
  const [level3, setLevel3] = useState<string>(dataset.columns[2].name);
  const [selectedItems, setSelectedItems] = useState<DynamicItem[]>([]);
  const [screen, setScreen] = useState<boolean>(false);

  //const selection = useConst(() => new Selection());
  const selection = useConst(
    () =>
      new Selection({
        onSelectionChanged: () => {
          setSelectedItems(selection.getSelection() as DynamicItem[]);
        },
      })
  );
  

  useEffect(() => {
    const { groups: generatedGroups, items: itemsFromMakeGroup } = makeGroups(dataset, level1 ?? "", level2 ?? "");
    setGroups(generatedGroups);
  }, [dataset, level1, level2, level3]);

  useEffect(() => {
    if (groups.length > 0) {
      const { items: extractedItems, columns: extractedColumns } = makeColumnAndItems(dataset, level1 ?? "", level2 ?? "", groups);
      setItems(extractedItems);
      setColumns(extractedColumns);
      selection.setItems(extractedItems, true);
    }
  }, [dataset, level1, level2, level3, groups]);

  useEffect(() => { //Ne znam da li je potrebno, u principu i nije, al nek stoji za dalju upotrebu
    const selectedRecordIds = selectedItems.map((item) => item.ct_bwprocessitemguid as string);
    dataset.setSelectedRecordIds(selectedRecordIds);
  }, [selection, dataset]);
  

  const openRecordForm = (item: DynamicItem) => {
    context.navigation.openForm({
      entityName: "ct_bw_procesitems", //Entity on which is called
      entityId: item.ct_bwprocessitemguid as string
    });
  };

  const onRenderCell = (nestingDepth?: number, item?: DynamicItem, itemIndex?: number, group?: IGroup): React.ReactNode => {
    //console.log("Selected items: ", selectedItems);

    const selectedRecordIds = selectedItems.map((item) => item.ct_bwprocessitemguid as string);
    dataset.setSelectedRecordIds(selectedRecordIds);
    //console.log("IDs: ", selectedRecordIds);

    //console.log("Dataset: ", dataset.getSelectedRecordIds());

    return item && typeof itemIndex === "number" && itemIndex > -1 ? (
      <div onDoubleClick={() => openRecordForm(item)}>
        <DetailsRow
          columns={columns}
          groupNestingDepth={nestingDepth}
          item={item}
          itemIndex={itemIndex}
          selection={selection}
          selectionMode={SelectionMode.multiple}
          group={group}
        />
      </div>
    ) : null;
  };

  return (
    <div className="container">
      {screen ? (
        <div className="selected-items-container">
          <h2>Selected Items</h2>
          {selectedItems.length > 0 ? (
            selectedItems.map((item) => (
              <div key={item.ct_name as string | number}>{item.ct_name}</div>
            ))
          ) : (
            <p>No items selected.</p>
          )}
          <button onClick={() => setScreen(false)}>Back</button>
        </div>
      ) : (
        <div className="grouped-list-container">
          <div className="header-div">
            <div className="columnSpacer"></div>
            <div className="header-columns">
              {columns.map((column) => (
                <div key={column.key} className="header-column">
                  {column.name}
                </div>
              ))}
            </div>
          </div>
          <SelectionZone selection={selection} selectionMode={SelectionMode.multiple}>
            <GroupedList
              items={items}
              onRenderCell={onRenderCell}
              selection={selection}
              selectionMode={SelectionMode.multiple}
              groups={groups}
              className="ms-GroupedList"
            />
          </SelectionZone>
        </div>
      )}
    </div>
  );
};