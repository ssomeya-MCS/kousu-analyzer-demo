import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components

# ==========================================
# ページ設定とカスタムCSS
# ==========================================
st.set_page_config(page_title="工数分析ダッシュボード Pro 3.0", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .metric-card { background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef; text-align: center; }
    .metric-value { font-size: 24px; font-weight: bold; color: #3498db; }
    .metric-label { font-size: 14px; color: #7f8c8d; }
    button[data-baseweb="tab"] { font-size: 16px !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("📊 工数データドリブン基盤 Pro 3.0 (デモ版)")
st.markdown("このアプリは、カレンダー入力UIの操作感テストと、Excelデータ取り込みによる分析ダッシュボードのデモです。")

# ==========================================
# データの読み込み関数
# ==========================================
@st.cache_data
def load_data(file):
    try:
        # 1行目をヘッダーとして読み込み
        df = pd.read_excel(file, sheet_name=0)
        
        if df.empty:
            return None, "Excelファイルにデータが存在しません。"

        # 列名の前後の空白を削除
        df.columns = df.columns.astype(str).str.strip().str.replace(' ', '')

        if '作業(リストから選択)' in df.columns:
            df.rename(columns={'作業(リストから選択)': '作業'}, inplace=True)

        if '日' in df.columns:
            df['日付'] = pd.to_datetime(df['日'], errors='coerce')
        else:
            return None, "「日」という列名が見つかりません。"

        df['年月'] = df['日付'].dt.strftime('%Y-%m')
            
        # ★実績時間の取得
        if '実績時間（分）' in df.columns:
            df['実績時間'] = pd.to_numeric(df['実績時間（分）'], errors='coerce').fillna(0) / 60.0
        else:
            df['実績時間'] = 0.0

        # ★予定時間の取得（追加）
        if '予定時間（分）' in df.columns:
            df['予定時間'] = pd.to_numeric(df['予定時間（分）'], errors='coerce').fillna(0) / 60.0
        elif '予定時間' in df.columns:
            df['予定時間'] = pd.to_numeric(df['予定時間'], errors='coerce').fillna(0)
        else:
            df['予定時間'] = 0.0
            
        # 処理件数の取得
        if '実績件数' in df.columns:
            df['処理件数'] = pd.to_numeric(df['実績件数'], errors='coerce').fillna(0)
        else:
            df['処理件数'] = 0
            
        # 欠損値の穴埋め
        for col in ['対象会社', '区分', '業務', '作業']:
            if col in df.columns:
                df[col] = df[col].fillna('未分類')
            else:
                df[col] = '未分類'

        df = df.dropna(subset=['日付'])
        
        if df.empty:
            return None, "有効なデータ行が1件もありません。"

        return df, None

    except Exception as e:
        return None, f"ファイルの読み込み中に予期せぬエラーが発生しました: {str(e)}"

# ==========================================
# サイドバー：ファイルアップロード & フィルタ
# ==========================================
st.sidebar.header("📁 データ読み込み")
uploaded_file = st.sidebar.file_uploader("実績Excelファイルをアップロード", type=["xlsx", "xls"])

df = None
if uploaded_file is not None:
    df, error_msg = load_data(uploaded_file)
    
    if error_msg:
        st.error("【読み込みエラー】")
        st.write(error_msg)
        df = None
    elif df is not None and not df.empty:
        st.sidebar.success("読み込み成功！")
        st.sidebar.header("🔍 分析条件フィルタ")

        # ★変更：年月を範囲指定スライダーに変更（初期値：6ヶ月前～最新月）
        months = sorted(list(df['年月'].dropna().unique()))
        if len(months) > 1:
            default_start = months[-6] if len(months) >= 6 else months[0]
            default_end = months[-1]
            start_month, end_month = st.sidebar.select_slider(
                "分析対象期間",
                options=months,
                value=(default_start, default_end)
            )
        elif len(months) == 1:
            st.sidebar.write("分析対象期間: ", months[0])
            start_month, end_month = months[0], months[0]
        else:
            start_month, end_month = None, None

        kbn_list = ["すべて"] + list(df['区分'].unique())
        sel_kbn = st.sidebar.selectbox("区分 (定例/非定例など)", kbn_list)

        gyomu_list = ["すべて"] + list(df['業務'].unique())
        sel_gyomu = st.sidebar.selectbox("業務カテゴリ", gyomu_list)

        # フィルタリング
        filtered_df = df.copy()
        if start_month and end_month:
            filtered_df = filtered_df[(filtered_df['年月'] >= start_month) & (filtered_df['年月'] <= end_month)]
        if sel_kbn != "すべて":
            filtered_df = filtered_df[filtered_df['区分'] == sel_kbn]
        if sel_gyomu != "すべて":
            filtered_df = filtered_df[filtered_df['業務'] == sel_gyomu]

# ==========================================
# メイン画面：タブ構成
# ==========================================
tab1, tab2, tab3 = st.tabs(["📅 カレンダー入力 (UIデモ)", "📈 分析ダッシュボード", "📄 カレントデータ & 差異申告"])

# ------------------------------------------
# タブ1: カレンダー入力UIモック
# ------------------------------------------
with tab1:
    st.info("💡 **【UIモックアップ】** ドラッグ＆ドロップによる入力操作のテスト画面です。（データは保存されません）")
    
    # ★変更：対象会社、業務、作業名を入力できるようにHTML側を改修
    calendar_html = """
    <!DOCTYPE html>
    <html>
    <head>
      <script src='https://cdn.jsdelivr.net/npm/fullcalendar@6.1.10/index.global.min.js'></script>
      <style>
        body { font-family: sans-serif; margin: 0; padding: 0; background: #fff; }
        #calendar { max-width: 100%; height: 600px; margin: 0 auto; }
        .modal { display: none; position: fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:2000; align-items:center; justify-content:center; }
        .modal-content { background: #fff; padding: 20px; border-radius: 8px; width: 340px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
        .modal input, .modal select { width: 100%; padding: 8px; margin: 4px 0 12px 0; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        label { font-size: 12px; font-weight: bold; color: #555; }
        .btn { background: #3498db; color: #fff; padding: 10px; width: 100%; border: none; border-radius: 4px; cursor: pointer; margin-top: 5px; font-weight: bold; }
        .btn-cancel { background: #95a5a6; }
      </style>
    </head>
    <body>
      <div id='calendar'></div>
      
      <!-- 簡易入力モーダル -->
      <div id="mockModal" class="modal">
        <div class="modal-content">
          <h3 id="modalTitle" style="margin-top:0; border-bottom:2px solid #3498db;">実績の入力</h3>
          
          <label>対象会社 (任意)</label>
          <input type="text" id="m-company" placeholder="例: 株式会社G">
          
          <label>区分</label>
          <select id="m-kbn"><option>定例</option><option>非定例</option></select>
          
          <label>業務 (カテゴリ)</label>
          <input type="text" id="m-biz" placeholder="例: プロジェクト">
          
          <label>作業名</label>
          <input type="text" id="m-task" placeholder="例: 月次資料作成">
          
          <label>処理件数</label>
          <input type="number" id="m-count" value="1">
          
          <button class="btn" onclick="saveMockEvent()">仮保存（デモ）</button>
          <button class="btn btn-cancel" onclick="closeModal()">キャンセル</button>
        </div>
      </div>

      <script>
        var calendar;
        var currentSelection = null;
        
        document.addEventListener('DOMContentLoaded', function() {
          var calendarEl = document.getElementById('calendar');
          calendar = new FullCalendar.Calendar(calendarEl, {
            initialView: 'timeGridWeek',
            slotMinTime: "08:00:00",
            slotMaxTime: "20:00:00",
            selectable: true,
            editable: true,
            headerToolbar: { left: 'prev,next today', center: 'title', right: 'timeGridWeek,timeGridDay' },
            events: [
              { title: '(定例) 朝礼', start: new Date(new Date().setHours(9,0,0,0)), end: new Date(new Date().setHours(9,30,0,0)), backgroundColor: '#3498db' },
              { title: '(非定例) 採用面接', start: new Date(new Date().setHours(11,0,0,0)), end: new Date(new Date().setHours(12,0,0,0)), backgroundColor: '#e74c3c' }
            ],
            select: function(info) {
              currentSelection = info;
              document.getElementById('m-company').value = '';
              document.getElementById('m-biz').value = '';
              document.getElementById('m-task').value = '';
              document.getElementById('mockModal').style.display = 'flex';
            },
            eventClick: function(info) {
              alert('【クリックイベント】\\n登録内容: ' + info.event.title + '\\n※本番環境ではここで詳細編集画面が開きます。');
            }
          });
          calendar.render();
        });

        function closeModal() {
          document.getElementById('mockModal').style.display = 'none';
          calendar.unselect();
        }

        function saveMockEvent() {
          var biz = document.getElementById('m-biz').value || '未設定';
          var task = document.getElementById('m-task').value || '未設定';
          var kbn = document.getElementById('m-kbn').value;
          var color = kbn === '定例' ? '#3498db' : '#e74c3c';
          
          if(currentSelection) {
            calendar.addEvent({
              title: '(' + kbn + ') ' + biz + ' - ' + task,
              start: currentSelection.startStr,
              end: currentSelection.endStr,
              backgroundColor: color
            });
          }
          closeModal();
        }
      </script>
    </body>
    </html>
    """
    components.html(calendar_html, height=680, scrolling=True)

# ------------------------------------------
# タブ2: 分析ダッシュボード
# ------------------------------------------
with tab2:
    if df is None:
        st.warning("👈 左側のサイドバーからExcelファイルを取り込むと、ここにダッシュボードが表示されます。")
    else:
        st.subheader("分析サマリ")
        
        total_hours = filtered_df['実績時間'].sum()
        total_plan = filtered_df['予定時間'].sum()
        total_count = filtered_df['処理件数'].sum()
        avg_time = total_hours / total_count if total_count > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.markdown(f'<div class="metric-card"><div class="metric-label">総実績工数 (予定: {total_plan:,.1f}h)</div><div class="metric-value">{total_hours:,.1f} h</div></div>', unsafe_allow_html=True)
        col2.markdown(f'<div class="metric-card"><div class="metric-label">総処理件数</div><div class="metric-value">{total_count:,.0f} 件</div></div>', unsafe_allow_html=True)
        col3.markdown(f'<div class="metric-card"><div class="metric-label">1件当たり平均時間</div><div class="metric-value">{avg_time:,.2f} h/件</div></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            st.write("##### 📊 業務別の工数シェア（円グラフ）")
            pie_data = filtered_df.groupby('業務')['実績時間'].sum().reset_index()
            pie_data = pie_data[pie_data['実績時間'] > 0]
            if not pie_data.empty:
                fig_pie = px.pie(pie_data, values='実績時間', names='業務', hole=0.4)
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.write("データがありません")
                
        with chart_col2:
            st.write("##### 📈 工数推移 ＆ 予定時間比較")
            # ★追加：色分け軸をラジオボタンで切り替え
            color_breakdown = st.radio("棒グラフの分類軸:", ["区分 (定例/非定例)", "業務分類"], horizontal=True)
            breakdown_col = '区分' if color_breakdown == "区分 (定例/非定例)" else '業務'

            bar_data = filtered_df.groupby(['年月', breakdown_col])['実績時間'].sum().reset_index()
            line_data = filtered_df.groupby(['年月'])['予定時間'].sum().reset_index()

            if not bar_data.empty:
                # ★追加：Plotly Graph Objects を使った 複合グラフ（棒＋折れ線）
                fig_combo = px.bar(bar_data, x='年月', y='実績時間', color=breakdown_col)
                
                fig_combo.add_trace(go.Scatter(
                    x=line_data['年月'], y=line_data['予定時間'],
                    mode='lines+markers', name='予定時間',
                    line=dict(color='#e74c3c', width=3),
                    marker=dict(size=8)
                ))
                
                fig_combo.update_layout(
                    margin=dict(t=0, b=0, l=0, r=0), 
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_combo, use_container_width=True)
            else:
                st.write("データがありません")

        st.write("##### 📋 詳細集計テーブル (1件当たり時間)")
        table_data = filtered_df.groupby(['対象会社', '業務', '作業']).agg(
            予定工数_h=('予定時間', 'sum'),
            実績工数_h=('実績時間', 'sum'),
            総処理件数=('処理件数', 'sum')
        ).reset_index()
        
        table_data['1件当たり時間_h'] = (table_data['実績工数_h'] / table_data['総処理件数']).round(2)
        table_data['1件当たり時間_h'] = table_data['1件当たり時間_h'].fillna(0)
        table_data = table_data.sort_values(by='実績工数_h', ascending=False)
        st.dataframe(table_data, use_container_width=True, hide_index=True)

# ------------------------------------------
# タブ3: カレントデータ & 差異申告
# ------------------------------------------
with tab3:
    if df is None:
        st.warning("👈 Excelファイルを取り込むと、ここに生データと差異申告フォームが表示されます。")
    else:
        # ★追加：左側に差異申告フォーム、右側にカレントデータテーブル
        col_left, col_right = st.columns([1, 2.5])

        with col_left:
            st.write("##### 📝 差異申告フォーム")
            st.write("予定と実績のズレを報告します。")
            with st.form("variance_form"):
                v_date = st.date_input("年月日")
                v_kbn = st.selectbox("分類 (必須)", ["定例", "非定例", "その他"])
                v_task = st.text_input("業務名 (必須)")
                
                c1, c2 = st.columns(2)
                v_plan = c1.number_input("予定 (h)", step=0.5, min_value=0.0)
                v_actual = c2.number_input("実績 (h)", step=0.5, min_value=0.0)
                
                v_reason = st.text_area("差異理由 (必須)", placeholder="例: 急なインシデント対応のため予定より2時間超過")
                
                submitted = st.form_submit_button("申告する", type="primary", use_container_width=True)
                
                if submitted:
                    if not v_task or not v_reason:
                        st.error("業務名と差異理由は入力必須です。")
                    else:
                        st.success("✅ 差異申告を受け付けました！（※デモ用のためデータは保存されません）")

        with col_right:
            st.write("##### 📄 カレントデータ (生データ)")
            st.write(f"全 {len(filtered_df):,} 件（フィルタ条件適用後）")
            display_cols = [c for c in ['日付', '曜日', '対象会社', '区分', '業務', '作業', '予定時間（分）', '実績時間（分）', '実績件数', '備考'] if c in df.columns]
            st.dataframe(filtered_df[display_cols], use_container_width=True, height=600)