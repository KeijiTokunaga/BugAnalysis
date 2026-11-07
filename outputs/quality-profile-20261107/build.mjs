
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const groups=[["プロジェクト共通",[["基本","プロジェクトID","P-001","他の台帳と共通の識別子"],["基本","プロジェクト名","受注管理システム刷新","正式名称"],["基本","対象業務・目的","受注から出荷までを管理","利用者が達成する目的"],["基本","主な利用者","営業担当・出荷担当","役割と利用条件"],["基本","対象範囲・対象外","受注・出荷。会計は対象外","評価する境界を明記"],["品質","優先する品質","データ保全・正確性","優先順位と理由"],["品質","品質目標と測定条件","応答時間の目標、負荷・環境・測定方法","数値・単位・対象を併記"],["利用","想定利用者数",null,"人数。対象・集計期間は根拠欄へ"],["利用","想定処理件数",null,"件数。処理の単位・期間は根拠欄へ"],["利用","ピーク負荷・利用時間帯","月末に集中、平日9〜18時","同時利用・繁忙期"],["構成","システム構成・主要技術","Webアプリ・DB・夜間バッチ","主要構成と技術の版"],["構成","外部連携・依存先","決済サービス・在庫API","連携内容と責任分界"],["構成","対応環境","OS・ブラウザ・端末の対象版","サポート対象"],["規則","コード量の計数定義","物理行、空行・コメント・生成物は除外","測定ツール・版・言語別集計"],["規則","要件・受け入れ条件の計数定義","要件ID単位。条件は独立に合否判定可能な単位","分割粒度・対象・版"],["規則","不具合1件の定義","同一不具合の重複票を統合","テストNG・失敗処理とは別計数"],["規則","重大度の定義","金銭・情報・業務停止などの影響で分類","重大度ごとの具体的判定基準"],["運用","監視・障害検知","エラー監視・処理件数照合","観測対象外も明記"],["運用","復旧・データ保護","バックアップ・復元・切戻し","復旧目標と確認方法"],["運用","台帳・判断責任者","仕様・不具合台帳URL、役割名","個人情報は必要最小限"]]],["機能別",[["識別","プロジェクトID","P-001","共通票と一致"],["識別","機能ID・名称","F-001 決済確定","機能ごとに票を複製"],["業務","重要シナリオ","支払確定後に注文を確定","開始条件と成功条件"],["業務","関連要件・受け入れ条件IDと版","REQ-001 v2 / AC-001","関連仕様へ参照"],["影響","失敗時の業務・金銭・安全への影響","二重請求、注文停止","具体的な結果と影響範囲"],["影響","情報の機密性・権限","個人情報あり、管理者操作あり","情報種類とアクセス境界"],["影響","取消・再実行・復旧の可否","取消可能、二重実行防止が必要","回避策と手作業の有無"],["利用","利用者・利用頻度・対象期間","購入者、注文ごと","利用機会の定義"],["利用","処理試行・成功・失敗の定義","決済確定要求を1試行と数える","再試行・重複・中断の扱い"],["難しさ","外部連携と障害時の動作","決済API。タイムアウト時に照合","依存先数だけで判断しない"],["難しさ","状態遷移・並行処理・整合性","支払と注文の状態を同期","競合・非同期・共有データ"],["検証","重要な異常系・境界条件","二重送信、通信断、金額境界","検証すべきリスク"],["観測","計測方法・対象外","APIログ。画面での操作断念は未計測","観測可能な範囲を明示"]]],["リリース別",[["識別","プロジェクトID","P-001","共通票と一致"],["識別","リリースID・対象版","R-001 / v1.0","リリースごとに票を複製"],["識別","対象機能ID・変更ID","F-001 / CHG-001","変更の対象範囲"],["変更","開発区分","既存改修","新規・改修・移行・混在"],["変更","変更内容・流用範囲","決済処理を改修、認証は流用","対象外も明記"],["規模","新規コード行数",null,"共通票の計数定義に従う"],["規模","変更コード行数",null,"追加・削除との重複計数ルールを明記"],["規模","削除コード行数",null,"単位：行。言語別なら票を分ける"],["規模","対象要件数",null,"対象版の要件数"],["規模","対象受け入れ条件数",null,"対象版の条件数"],["要件","未決事項・曖昧な仕様","例外時の取消条件が未合意","判断者と解決予定"],["要件","合意後の要件変更","結合テスト開始後に条件追加","内容・時期・影響・関連ID"],["体制","役割別体制・業務技術経験","開発3名、業務経験者1名","必要な範囲で役割別に記録"],["体制","交代・引継ぎ・外部依存","設計担当交代、引継ぎ資料あり","情報共有に影響した条件"],["日程","開発・検証期間と計画差","検証期間を10日から5日に変更","予定と実績、変更理由"],["制約","期限・環境・データの制約","期限固定、本番相当データなし","選択肢に影響した条件"],["検証","検証工程・方法・自動化範囲","単体自動化、結合は手動","レビューの範囲も含む"],["検証","検証済み受け入れ条件数",null,"実行回数ではなく重複を除いた条件数"],["検証","未検証・対象外の重要シナリオ","通信断は環境不足で未検証","対象外と未実施を区別"],["検証","検証環境と本番との差","外部連携はスタブで確認","データ・負荷・権限の差"],["観測","観測開始日・終了日","本番公開から30日など","実際の日付と期間"],["観測","利用・失敗処理の観測記録への参照","ログ集計URL / OBS-001","機能・シナリオ別の分母分子へ参照"],["観測","観測状態","未計測","計測済み・利用機会なし・未計測"],["判断","既知の問題・残余リスク","未検証事項、回避策、判断者","リリース時点の判断を保存"]]]];
const out="/Users/keizi/orca/projects/BugAnalysis/outputs/quality-profile-20261107";
const wb=Workbook.create();
for(const [name,items] of groups){
 const s=wb.worksheets.add(name); s.showGridLines=false;
 s.getRange('A2').values=[[name+'｜品質評価のための特性入力票']];
 s.getRange('A3').values=[['黄色が入力欄。各シートを対象ごとに複製し、過去版を上書きせず保存。']];
 s.getRange('A4').values=[['空欄は未入力。数値の不明は空欄のまま状態を選択。0は確認済みのゼロのみ。']];
 const headers=['分類','入力項目','入力値','状態','入力例','説明・単位','根拠・参照先','基準日','記録者'];
 s.getRange('A6:I6').values=[headers];
 const rows=items.map(([cat,label,example,note])=>[cat,label,null,'未確認',example,note,null,null,null]);
 s.getRange('A7:I'+(6+rows.length)).values=rows;
 const end=6+rows.length;
 s.getRange('A2:I'+end).format.font={name:'Arial',size:11,color:'#243247'};
 s.getRange('A2').format.font={name:'Arial',size:15,bold:true};
 s.getRange('A6:I6').format={fill:'#263D58',font:{color:'#FFFFFF',bold:true},rowHeight:30,wrapText:true};
 s.getRange('A7:I'+end).format.rowHeight=62;
 s.getRange('A7:I'+end).format.wrapText=true;
 s.getRange('A7:I'+end).format.verticalAlignment='center';
 for(const col of ['C','D','G','H','I']) s.getRange(col+'7:'+col+end).format.fill='#FFF2CC';
 const widths={A:12,B:34,C:40,D:15,E:42,F:43,G:30,H:15,I:15};
 for(const [col,width] of Object.entries(widths)) s.getRange(col+'1:'+col+end).format.columnWidth=width;
 s.getRange('H7:H'+end).setNumberFormat('yyyy-mm-dd');
 s.getRange('D7:D'+end).dataValidation={rule:{type:'list',values:['未確認','確定','推定','不明','対象外']}};
 for(let i=0;i<items.length;i++) {
  if(items[i][2]===null){
   const r=i+7; s.getRange('C'+r).setNumberFormat('#,##0');
   s.dataValidations.add({range:'C'+r,rule:{type:'whole',operator:'greaterThanOrEqual',formula1:0}});
  }
 }
 s.freezePanes.freezeRows(6);
 assert.equal(rows.length,items.length);
 assert(rows.every(r=>r.length===9 && r[2]===null));
 const preview=await wb.render({sheetName:name,range:'A6:F12',scale:1,format:'png'});
 await fs.writeFile(out+'/'+name+'.png',new Uint8Array(await preview.arrayBuffer()));
}
wb.recalculate();
console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:1500})).ndjson);
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(out+'/プロジェクト特性入力テンプレート.xlsx');
console.log('Saved: '+out+'/プロジェクト特性入力テンプレート.xlsx');
