import PySimpleGUI as sg
import sys
from pathlib import Path
import logging


VERSION = '1.0.0'


logger = logging.getLogger(__name__)


class MainWindow:
    TITLE = f'Delrow ver{VERSION}'
    ENCODING = 'utf-8'
    LISTBOX_Y = 30

    def __init__(self, files: list[Path]) -> None:
        logger.debug('MainWindow instance initialize start.')
        self.cbFilePath = sg.Combo(
            files, default_value=files[0], key="cbFP",
            expand_x=True, readonly=True, enable_events=True
        )
        self.lbFileRows = sg.Listbox(
            ['読み込み中です...'], select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED,
            size=(80, self.LISTBOX_Y), expand_x=True, expand_y=True,
            horizontal_scroll=True,
            enable_events=True, key='lbFR'
        )
        self.tSelectedRowsInfo = sg.Text('-')

        # ##### レイアウト設定
        self.layout = [[sg.Text('表示中のテキストファイル：'), self.cbFilePath],
        [self.lbFileRows],
        [sg.Button('削除(D)', key='-削除-', size=(30,2)), self.tSelectedRowsInfo]
        ]

        self.window = sg.Window(
            self.TITLE, self.layout,
            resizable=True, return_keyboard_events=True
        )
        self.window.finalize()
        logger.debug('Window Opened.')
        self.nowfile = files[0]
        self._readFile(files[0])

    def mainloop(self) -> None:
        """ウインドウイベントループ
        """
        while True:
            event, values = self.window.read()  #type: ignore
            logger.debug(f'event={event} : values={values}')
            if event == sg.WINDOW_CLOSED:
                break
            elif event == 'lbFR':
                self._printSelectedRowsInfo()
            elif event == 'cbFP':
                self._changeFile(Path(values['cbFP']))
            elif event == '-削除-' or event == 'd':
                # dキーを入力しても削除処理に移る
                self._deleteRows(Path(values['cbFP']))

        self.window.close()
        logger.debug('Window Closed.')

    def _readFile(self, filepath: Path) -> None:
        """ファイルの内容を読み込み、リストボックスに表示する

        Args:
            filepath (Path): ファイルパス（コンボボックスに表示されているものを渡すこと）
        """
        logger.debug(f'_readFile({filepath}) started.')
        with open(filepath, 'r', encoding=self.ENCODING) as f:
            data: list[str] = []
            self.lbFileRows.update(
                values=['読み込み中です...'], set_to_index=0, disabled=True)
            self.window.refresh()
            for r in f:
                data.append(r)
            self.lbFileRows.update(values=data, set_to_index=0, disabled=False)
            self.tSelectedRowsInfo.update('-')
            self.window.refresh()

    def _printSelectedRowsInfo(self) -> None:
        """選択行案内表示の更新
        """
        selrow = self._getSelectedRows()
        val: str
        if selrow is None:
            val = '-'
        elif selrow == (-1,-1):
            val = '不正な選択です。連続した行を選択してください。'
        else:
            start = selrow[0] + 1
            end = selrow[1] + 1
            val = f'{start}行～{end}行を選択中。'

        self.tSelectedRowsInfo.update(value=val)  #type: ignore

    def _getSelectedRows(self) -> tuple[int,int] | None:
        """選択行を取得

        Returns:
            tuple[int,int] | None: 選択行の(開始行,終了行) : 不正な選択の場合は(-1,-1) : 未選択（あるのか？）の場合はNone
        """
        selected = self.lbFileRows.get_indexes()
        if len(selected) == 0:
            return None
        else:
            start = selected[0]
            k = start + 1
            for n in selected[1:]:
                if k != n:
                    return -1, -1
                else:
                    k += 1
            return start, k - 1

    def _changeFile(self, newfilepath: Path) -> None:
        """コンボボックス変更イベント 表示ファイルを切り替える

        Args:
            newfilepath (Path): コンボボックスの内容
        """
        if sg.popup_ok_cancel(
            f'表示ファイルを変更します\n{newfilepath.name}\n\nよろしいですか？',
            title='ファイル変更の確認'
        ) != 'OK':
            self.cbFilePath.update(value=self.nowfile)
            return
        self._readFile(newfilepath)

    def _deleteRows(self, filepath: Path) -> None:
        """削除処理の実行

        Args:
            filepath (Path): コンボボックスの内容
        """
        selrow = self._getSelectedRows()
        if selrow is None or selrow == (-1,-1):
            sg.popup_error('行が正しく選択されていません。', title='エラー - Delrow')
            return
        start, end = selrow
        if sg.popup_ok_cancel(
            (f'ファイル {filepath.name} の {start+1}行目～{end+1}行目を削除します。\n)'
             'よろしいですか？'),
            title='削除の確認'
        ) != 'OK':
            return

        data: list[str] = []
        with open(filepath, 'r', encoding=self.ENCODING) as rf:
            data = rf.readlines()
        logger.info((
            '対象ファイル %s\n'
            '↓↓削除するデータ↓↓\n%s\n↑↑↑↑↑↑↑↑↑↑↑')
                    , filepath, ''.join(data[start:end+1])[:-1])
        newdata = data[:start] + data[end+1:]
        try:
            with open(filepath, 'w', encoding=self.ENCODING) as wf:
                wf.writelines(newdata)
        except PermissionError:
            errormes = f'ファイル「{filepath}」は開かれているため変更できません。'
            sg.popup_error(errormes, title='エラー - Delrow')
            logger.error('PermissionError: ' + errormes)
            return
        sg.popup('削除しました。', title='削除完了')
        self._readFile(filepath)


def _fileExistCheck(fp: Path, raise_exc: bool = True) -> bool:
    b = fp.exists()
    if raise_exc and not b:
        raise FileNotFoundError(f'ファイル「{fp}」は有りません。')
    return b


if __name__ == '__main__':
    # ##### ロガー設定
    logging.basicConfig(
        format='%(asctime)s %(levelname)s in %(name)s: %(message)s',
        level=logging.DEBUG)
    fileloghandle = logging.FileHandler('delrow.log','a', encoding='utf-8')
    fileloghandle.setLevel(logging.INFO)
    logformatter = logging.Formatter(
        '%(asctime)s %(levelname)s in %(name)s: %(message)s'
    )
    fileloghandle.setFormatter(logformatter)
    logger.addHandler(fileloghandle)
    

    try:
        if len(sys.argv) < 2:
            raise RuntimeError('読み込むテキストファイルをD&Dしてください')

        files: list[Path] = []
        for p in sys.argv[1:]:
            _fileExistCheck(Path(p))
            files.append(Path(p))

        mwin = MainWindow(files)
        mwin.mainloop()

    except Exception as e:
        sg.popup_error(f'{e.__class__.__name__} :{e}', title='エラー - Delrow')
        logger.critical(e, exc_info=True)
        sys.exit(1)
