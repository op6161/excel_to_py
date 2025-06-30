from .excel_module import DataHandler
from openpyxl.cell.cell import MergedCell

class ExtendHandler(DataHandler):
    def __init__(self, excel_data, sheet_name=None, fill_dummy_value=None, dummy_value="MERGED"):
        super().__init__(excel_data, sheet_name)
        self.fill_dummy_value = fill_dummy_value
        self.dummy_value = dummy_value

    def _get_cell_value_from_merged_range(self, sheet, cell):
        """
        セルがMergedCellの場合、その結合範囲の先頭セルの値を返します。
        それ以外の場合は、セルの値をそのまま返します。
        """
        if isinstance(cell, MergedCell):
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return sheet[merged_range.start_cell.coordinate].value
        return cell.value

    def _get_raw_cell_value(self, sheet, cell):
        """
        DataHandlerの_get_raw_cell_valueをオーバーライドし、
        結合されたセルとオプションのダミー値を処理します。
        """
        value = self._get_cell_value_from_merged_range(sheet, cell)
        
        # 병합된 셀 처리 후에도 값이 None이고, 더미 값 채우기가 활성화된 경우
        if value is None and self.fill_dummy_value:
            return self.dummy_value
        
        return value

    def _get_cell_value_with_merged_fill(self, sheet, cell):
        """
        병합된 셀의 값을 가져오는 내부 유틸리티 함수 (한국어 설명)
        MergedCellであれば、その結合された範囲の左上隅のセルの値を返す。
        それ以外であれば、セルの値をそのまま返す。
        self.fill_merged_cells이 True일 경우, 병합된 셀의 빈 공간을 첫 번째 셀 값으로 채웁니다.
        """
        if isinstance(cell, MergedCell):
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return sheet[merged_range.start_cell.coordinate].value
        
        # fill_merged_cells가 True이고, 셀 값이 None 또는 빈 문자열이면 dummy_value로 채움
        # 단, MergedCell이 아니면서 원래 비어있던 셀은 dummy_value로 채우지 않도록 조건 추가
        # 즉, 병합된 셀이 'None'이 되는 경우만 해당 더미 값을 적용하도록 합니다.
        if self.fill_merged_cells and cell.value is None: # None만 처리하도록 명확히
            # MergedCell이 아닌데도 None인 경우 (원래 빈 셀)는 채우지 않음
            # 이 로직은 `_get_merged_cell_value`가 처리하지 못한 'None'을 위한 것이며,
            # MergedCell이지만 첫 번째 셀이 이미 None인 경우도 처리할 수 있습니다.
            return self.dummy_value if self.dummy_value is not None else cell.value
        
        return cell.value

    def _get_cell_value_with_dummy_fill(self, sheet, cell):
        """
        _get_cell_value_with_merged_fill을 호출하고, 그 결과가 None일 때 dummy_value로 채우는 함수.
        이 함수가 최종적으로 외부로 값을 반환할 때 사용됩니다.
        """
        value = self._get_cell_value_with_merged_fill(sheet, cell)
        if value is None and self.fill_merged_cells:
            return self.dummy_value
        return value

    def get_block(self, sheet_name=None):
        """

        """
        solved_sheet_name = self._resolve_sheet_name(sheet_name)
        blocks = find_data_blocks(self.get_sheet(solved_sheet_name),get_block_data=True)
        return blocks
    
    def get_block_points(self, sheet_name=None):
        """

        """
        solved_sheet_name = self._resolve_sheet_name(sheet_name)
        blocks = find_data_blocks(self.get_sheet(solved_sheet_name), get_block_data=False)
        return blocks

def find_data_blocks(sheet_data, include_diagonals=True, get_block_data = True):
    """
    Excelシートデータから隣接しているセルの集合（ブロック）を探します。
    
    Args:
        sheet_data (list of list): リスト化できるExcelシートデータ
        include_diagonals (bool): 対角のセルを隣接として演算するかないか。
        get_block_data (bool): 結果値にブロックの値包含可否

    Returns:
        list of dict: 各ブロックのデータ、開始/終了/列情報を含むディクショナリーリスト。
              例: [{"data": [[...]], "start_row": 0, "start_col": 0, "end_row": 1, "end_col": 2}, ...]
    """
    if not sheet_data:
        raise ValueError("空のシートです。")

    # 計算のため2Dリストに変換
    grid = [list(row) for row in sheet_data]
    
    rows = len(grid)
    cols = len(grid[0]) if rows > 0 else 0
    visited = [[False for _ in range(cols)] for _ in range(rows)]
    
    found_blocks = []

    # 探索する方向定義
    directions = [(0, 1), (0, -1), (1, 0), (-1, 0)] # ↓↑→←
    if include_diagonals:
        directions.extend([(1, 1), (1, -1), (-1, 1), (-1, -1)]) # ↘↗↙↖

    # dfsアルゴリズムを使用
    def dfs(r, c, current_block_coords):
        # チェック・訪問の有無・値(None)をチェック
        if not (0 <= r < rows and 0 <= c < cols) or visited[r][c] or grid[r][c] is None:
            return None

        visited[r][c] = True
        current_block_coords.append((r, c))

        # 定義された方向の隣接したセルの値の有無を確認
        for dr, dc in directions:
            dfs(r + dr, c + dc, current_block_coords)

    for r in range(rows):
        for c in range(cols):
            # 訪問しなかった、Noneではない新しいセルの確認
            if grid[r][c] is not None and not visited[r][c]:
                current_block_coords = []
                dfs(r, c, current_block_coords)
                
                # ブロックのサイズ計算
                min_r = rows
                max_r = -1
                min_c = cols
                max_c = -1
                
                for row_idx, col_idx in current_block_coords:
                    min_r = min(min_r, row_idx)
                    max_r = max(max_r, row_idx)
                    min_c = min(min_c, col_idx)
                    max_c = max(max_c, col_idx)
                        
                block = {"start_row": min_r,
                         "start_col": min_c,
                         "end_row": max_r,
                         "end_col": max_c}
                
                if get_block_data:
                    block_data = []
                    for row_idx in range(min_r, max_r + 1):
                        row_values = []
                        for col_idx in range(min_c, max_c + 1):
                            row_values.append(grid[row_idx][col_idx])
                        block_data.append(row_values)
                    block["data"] = block_data

                found_blocks.append(block)
                
    return found_blocks
