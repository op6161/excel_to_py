from .excel_module import DataHandler

class ExtendHandler(DataHandler):
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
