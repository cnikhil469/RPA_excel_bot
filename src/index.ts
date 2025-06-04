import { mouse, keyboard, Key, sleep } from "@nut-tree-fork/nut-js";

mouse.config.autoDelayMs = 100;
keyboard.config.autoDelayMs = 100;

async function openExcel() {
  await keyboard.pressKey(Key.LeftWin);
  await keyboard.pressKey(Key.R);
  await keyboard.releaseKey(Key.R);
  await keyboard.releaseKey(Key.LeftWin);

  await keyboard.type("excel");
  await keyboard.pressKey(Key.Enter);

  await sleep(2000);
}

async function createNewWorkbook() {
  await keyboard.pressKey(Key.LeftControl);
  await keyboard.pressKey(Key.N);
  await keyboard.releaseKey(Key.N);
  await keyboard.releaseKey(Key.LeftControl);
  await keyboard.pressKey(Key.Enter);
  await keyboard.releaseKey(Key.Enter);
  await keyboard.pressKey(Key.Enter);
  await keyboard.releaseKey(Key.Enter);
  await sleep(1000);
}

async function addDataToExcel() {
  const data = [
    ["Name", "Age", "City"],
    ["Nikhil", "25", "New York"],
    ["Rahul", "25", "Los Angeles"],
    ["Raj", "35", "Chicago"],
  ];

  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      await keyboard.type(data[i][j]);

      if (j < data[i].length - 1) {
        await keyboard.pressKey(Key.Tab);
      }
    }

    if (i < data.length - 1) {
      await keyboard.pressKey(Key.Enter);
    }
  }
}

async function saveWorkbook() {
  await keyboard.pressKey(Key.LeftControl);
  await keyboard.pressKey(Key.S);
  await keyboard.releaseKey(Key.S);
  await keyboard.releaseKey(Key.LeftControl);

  await sleep(1000);

  await keyboard.type("automated_data.xlsx");

  await keyboard.pressKey(Key.Enter);
}

async function main() {
  try {
    console.log("Now we are going to automate Excel!");

    await openExcel();
    await createNewWorkbook();
    await addDataToExcel();
    await saveWorkbook();

    console.log("Excel automation completed successfully!");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

main();
