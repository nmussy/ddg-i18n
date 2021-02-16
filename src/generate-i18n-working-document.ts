import {exec} from 'child_process';
import {config} from 'dotenv';
import {Workbook, Worksheet} from 'exceljs';
import {
  copyFile,
  existsSync,
  mkdirp,
  mkdirpSync,
  readdir,
  readFile,
  remove,
} from 'fs-extra';
import {glob} from 'glob';
import gm, {Dimensions} from 'gm';
import {basename, dirname, join} from 'path';
import naturalCompare from 'string-natural-compare';

config();

['PERFECT_PATH', 'DGF_TXM_CONVERT_PATH'].map((key) => {
  if (!process.env[key]) throw new Error(`Missing ${key} env variable`);
});

const {PERFECT_PATH, DGF_TXM_CONVERT_PATH} = process.env;
const OUTPUT_PATH = join(__dirname, '../out');
if (!existsSync(OUTPUT_PATH)) {
  mkdirpSync(OUTPUT_PATH);
}

const CELL_DIMENSION = {
  width: 64,
  height: 20,
};
const IMAGE_CELL_SIZE = {
  width: CELL_DIMENSION.width * 4,
  height: CELL_DIMENSION.height,
};

interface AddCellParams {
  sheet: Worksheet;
  rowIndex: number;
  cellIndex: number;
  length?: number;
  text?: string;
}

const addCell = ({
  sheet,
  rowIndex,
  cellIndex,
  length = 1,
  text = '',
}: AddCellParams): number => {
  sheet.mergeCells(rowIndex, cellIndex, rowIndex, cellIndex + length - 1);
  sheet.getCell(rowIndex, cellIndex).value = text;

  return cellIndex + length;
};

const generateSheet = async (workbook: Workbook, path: string) => {
  const sheet = workbook.addWorksheet(
    basename(path)
      .replace(/^_/, '')
      .replace(/\.dat$/, ''),
  );
  let rowIndex = 1;
  [
    {text: 'File name', length: 3},
    {text: 'Input texture', length: 4},
    {text: 'Kanji', length: 2},
    {text: 'English', length: 2},
    {text: 'Output texture', length: 4},
  ].reduce(
    (cellIndex, params) =>
      addCell({
        sheet,
        rowIndex,
        cellIndex,
        ...params,
      }),
    1,
  );

  const files = (await readdir(path)).sort(naturalCompare);
  for (const file of files) {
    ++rowIndex;
    const imagePath = join(path, file);
    const size = await new Promise<Dimensions>((resolve, reject) =>
      gm(imagePath).size((err, size) => (err ? reject(err) : resolve(size))),
    );

    const [ratioWidth, ratioHeight] = ['width', 'height'].map(
      (prop) => size[prop] / IMAGE_CELL_SIZE[prop],
    );
    const horizontalImage = ratioWidth > ratioHeight;

    const ext = {
      width: horizontalImage ? IMAGE_CELL_SIZE.width : size.width / ratioHeight,
      height: horizontalImage
        ? size.height / ratioWidth
        : IMAGE_CELL_SIZE.height,
    };

    const imgId = workbook.addImage({
      buffer: await readFile(imagePath),
      extension: 'png',
    });

    [
      {text: file.replace(/\.png$/, ''), length: 3},
      {length: 4},
      {length: 2},
      {length: 2},
      {length: 4},
    ].reduce(
      (cellIndex, params) =>
        addCell({
          sheet,
          rowIndex,
          cellIndex,
          ...params,
        }),
      1,
    );

    sheet.addImage(imgId, {
      tl: {col: 3, row: rowIndex - 1},
      editAs: undefined,
      ext,
    });
  }
};

(async () => {
  const cwd = join(PERFECT_PATH, 'cddata');
  const datFiles = await new Promise<string[]>((resolve, reject) =>
    glob(
      join('**/*.dat'),
      {
        ignore: [
          ...[
            'mapdata',
            '3d',
            'bg',
            'imagetop',
            'mapanim',
            'se',
            'taikou',
            'test',
            'tref',
            'trm',
            'trm2',
          ].map((dir) => `${dir}/**`),
          ...[
            'title_vram',
            'menu_vram_us',
            'menu_vram_fj',
            'com_vram',
            'testmap',
            'lect_vram',
            'lect_mem',
            'end_vram',
          ].map((file) => `**/${file}.dat`),
        ],
        cwd,
      },
      (err, matches) => (err ? reject(err) : resolve(matches)),
    ),
  );

  const workbook = new Workbook();
  for (const datRelativePath of datFiles.reverse()) {
    const _datPath = join(cwd, datRelativePath);
    console.time(`extract: ${basename(_datPath)}`);

    const tmpDir = join(
      dirname(_datPath),
      `_${basename(_datPath).replace(/\.dat$/, '')}`,
    );
    await mkdirp(tmpDir);

    const tmpDatPath = join(tmpDir, basename(_datPath));

    await copyFile(_datPath, tmpDatPath);

    const windowDatPath = tmpDatPath
      .replace(/^\/mnt\/(\w+)/, '$1:')
      .replace(/\//g, '\\\\');

    try {
      await new Promise((resolve, reject) =>
        exec(
          `${DGF_TXM_CONVERT_PATH} convert-dat ${windowDatPath}`,
          (err, out) => (err ? reject(err) : resolve(out)),
        ),
      );
    } catch (err) {
      console.error(err);
      return;
    } finally {
      await remove(tmpDatPath);
    }

    console.timeEnd(`extract: ${basename(_datPath)}`);

    console.time(`xlsx: ${basename(_datPath)}`);
    await generateSheet(workbook, tmpDir);
    console.timeEnd(`xlsx: ${basename(_datPath)}`);
  }

  await workbook.xlsx.writeFile(
    join(OUTPUT_PATH, `${new Date().getTime()}.xlsx`),
  );
})();
