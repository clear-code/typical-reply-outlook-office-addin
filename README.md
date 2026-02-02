# typical-reply-outlook-office-addin

## 概要

定型文での返信を提供するOutlook向けアドインです。
Officeアドインプラットフォームで作成されています。

例えば、以下のメールを受信したとします。

```
件名:
  例の件について

本文:
  お疲れ様です。
  例の件について以下の案を考えてみたのですが、いかがでしょうか？
  http://...
```

このとき、以下のいずれかの手順ですぐに定型文で返信することができます。

* メールを表示中のリボンのアドイン->「TypicalReply」グループから、「いいね！」を選択
  （ボタンが一つしかない場合はTypicalReplyボタンのみが表示される。）
  !["リボン"](./documents/images/ribbon.PNG "リボン")
* 閲覧ウィンドウ（ペイン）のアドイン->「TypicalReply」グループから、「いいね！」を選択
（ボタンが一つしかない場合はTypicalReplyボタンのみが表示される。）
  !["閲覧ウィンドウ（ペイン）"](./documents/images/message-reading-pain.PNG "閲覧ウィンドウ（ペイン）")

```
Subject:
  [[いいね！]]: Re: 例の件について

Body:
  いいね！
  
  > -----Original Message-----
  > お疲れ様です。
  > 例の件について以下の案を考えてみたのですが、いかがでしょうか？
  > http://...
```

返信内容については、設定ファイルによる設定変更が可能です。

また、新しいOutlook、Outlook on the webでは閲覧ウィンドウの操作のカスタマイズでアドインをメッセージリーディングペインに常に表示するようにすることも可能です。

!["閲覧ウィンドウの操作のカスタマイズ"](./documents/images/reading-window-custom.PNG "閲覧ウィンドウの操作のカスタマイズ")

!["ボタンが追加された閲覧ウィンドウ"](./documents/images/customized-reading-window-operations.PNG "ボタンが追加された閲覧ウィンドウ")

## 設定ファイルによる設定

設定ファイルは以下の箇所に存在します。

`C:\Program Files\TypicalReply\contents\Configs\TypicalReplyConfig.json`

`C:\Program Files\TypicalReply`はインストール先フォルダーです。

このファイルに後述のJSONでの設定を記載します。

## 設定項目

設定は以下のようなJSON形式で指定します。

```
{
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "ButtonConfigList": [
                {
                    "Id": "msgReadTypicalReply",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        },
        {
            "Culture": "en-US",
            "ButtonConfigList": [
                {
                    "Id": "msgReadTypicalReply",
                    "SubjectPrefix": "[[Like!]]:",
                    "Body": "Like!",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        }
    ]
}
```

TypiclReplyConfig: 設定のルート

| 設定名     | 型             | 必須 | 省略時のデフォルト | 概要                                                                                                                         |
| ---------- | -------------- | ---- | ------------------ | ---------------------------------------------------------------------------------------------------------------------------- |
| ConfigList | Configのリスト | yes  | -                  | 各言語ごとの設定（Config）のリスト                                                                                           |

Config: 各言語ごとの設定

| 設定名                    | 型                   | 必須 | 省略時のデフォルト   | 概要                                                                                                                                                                     | 例                         |
| ------------------------- | -------------------- | ---- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------- |
| Culture                   | 文字列               | no   | null                 | 対象となるカルチャ。<br>ロケールなしの言語のみを指定することも可能です。<br>現在のカルチャにマッチするCultureがない場合、Cultureの値に関わらず先頭のConfigを使用します。 | `"ja-JP"`、`"ja"`          |
| ButtonConfigList          | ButtonConfigのリスト | yes  | -                    | 定型返信ボタン設定のリスト                                                                                                                                               | -                          |



ButtonConfig: 定型返信ボタン設定。返信内容や返信先等の設定を行う。
| 設定名         | 型             | 必須 | 省略時のデフォルト   | 概要                                                                                                                                                                                        | 例                                        |
| -------------- | -------------- | ---- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------- |
| Id             | 文字列         | yes  | -                    | ボタンのID。ButtonConfigList内で重複不可。                                                                                                                                                  | `"LikeId"`                                |
| SubjectPrefix  | 文字列         | no   | null                 | 件名の先頭に挿入する文言                                                                                                                                                                    | `"[[いいね]]"`                            |
| Subject        | 文字列         | no   | 返信のデフォルト件名 | 件名                                                                                                                                                                                        | `"報告"`                                  |
| Body           | 文字列         | no   | null                 | 本文                                                                                                                                                                                        | `"いいね"`                                |
| Recipients     | 文字列のリスト | no   | 送信先なし           | 送信先。<br>`["blank"]`: 送信先なし<br> `["all"]`: 全員に返信<br>`["sender"]`: 送信者にだけ返信<br>その他の文字列リスト: 指定のアドレスに返信                                               | `["test@test.co.jp", "test2@test.co.jp"]` |
| QuoteType      | boolean        | no   | false                | 元の文言を引用するかどうか。 <br> `true`: 引用する<br>`false`: 引用しない                                                                                                                   | `true`                                    |
| AllowedDomains | 文字列のリスト | no   | 全て許可             | 送信を許可するドメインリスト。このドメイン以外が含まれている場合、返信用メールの作成、送信は行わない。<br>`["*"]`: 全て許可する<br>その他の文字列リスト: 指定したドメインのみ送信を許可する | `["test.co.jp", "test2.co.jp"]`           |
| ForwardType    | 文字列         | no   | 添付しない           | 元のメールを添付するかどうか。<br>`attachment`: 添付する                                                                                                                                    | `attachment`                              |

## 新しい設定追加の例

「最高！」というボタンを追加する方法を考えます。

ボタンの表示設定はマニフェストファイル（`TypicalReply.manifest.json`）で行います。

マニフェストファイルを開きます。

まずは、追加したいボタンのテキストを指定します。

現在の設定が以下のようになっているとします。

```
        <bt:ShortStrings>
          <bt:String id="CommandsGroup.Label" DefaultValue="Typical"/>
          <bt:String id="TypicalReplyButton.Label" DefaultValue="Like!">
            <bt:Override Locale="en-US" Value="Like!"/>
            <bt:Override Locale="ja-JP" Value="いいね!"/>
          </bt:String>
          <bt:String id="TypicalReplyButton.SupertipTitle" DefaultValue="Like!">
            <bt:Override Locale="en-US" Value="Like!"/>
            <bt:Override Locale="ja-JP" Value="いいね!"/>
          </bt:String>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TypicalReplyButton.SupertipText" DefaultValue="I like this message.">
            <bt:Override Locale="en-US" Value="I like this message."/>
            <bt:Override Locale="ja-JP" Value="このメッセージが気に入りました。"/>
          </bt:String>
        </bt:LongStrings>
      </Resources>
```

追加したいボタンのIDを`NewButton`とします。
`<bt:ShortStrings>`タグと`<bt:LongStrings>`タグに以下のようにボタンのラベル（`NewButton.Label`）、ツールチップのタイトル（`NewButton.SupertipTitle`）、ツールチップのテキスト（`NewButton.SupertipText`）を追加します。

`<bt:String>`タグは、125文字以下の場合は`<bt:ShortStrings>`タグ配下に、250文字以下の場合は`<bt:LongStrings>`タグ配下に定義します。
また、`id`は32文字以下である必要があります。

```
        <bt:ShortStrings>
          <bt:String id="CommandsGroup.Label" DefaultValue="Typical"/>
          <bt:String id="TypicalReplyButton.Label" DefaultValue="Like!">
            <bt:Override Locale="en-US" Value="Like!"/>
            <bt:Override Locale="ja-JP" Value="いいね!"/>
          </bt:String>
          <bt:String id="TypicalReplyButton.SupertipTitle" DefaultValue="Like!">
            <bt:Override Locale="en-US" Value="Like!"/>
            <bt:Override Locale="ja-JP" Value="いいね!"/>
          </bt:String>
          <bt:String id="NewButton.Label" DefaultValue="Awesome!">
            <bt:Override Locale="en-US" Value="Awesome!"/>
            <bt:Override Locale="ja-JP" Value="最高！"/>
          </bt:String>
          <bt:String id="NewButton.SupertipTitle" DefaultValue="Awesome!">
            <bt:Override Locale="en-US" Value="Awesome!"/>
            <bt:Override Locale="ja-JP" Value="最高！"/>
          </bt:String>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TypicalReplyButton.SupertipText" DefaultValue="I like this message.">
            <bt:Override Locale="en-US" Value="I like this message."/>
            <bt:Override Locale="ja-JP" Value="このメッセージが気に入りました。"/>
          </bt:String>
          <bt:String id="NewButton.SupertipText" DefaultValue="This message is Awesome!">
            <bt:Override Locale="en-US" Value="This message is Awesome!"/>
            <bt:Override Locale="ja-JP" Value="このメッセージは最高！"/>
          </bt:String>
        </bt:LongStrings>
      </Resources>
```

次に、`<ExtensionPoint xsi:type="MessageReadCommandSurface">` -> `<OfficeTab id="msgReadTabDefault">` -> `<Group id="msgReadCmdGroup">`配下に`<Control>`タグを追加します。

現在の設定が以下のようになっているとします。

```
            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="msgReadTabDefault">
                <Group id="msgReadCmdGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="msgReadTypicalReply">
                    <Label resid="TypicalReplyButton.Label"/>
                    <Supertip>
                      <Title resid="TypicalReplyButton.SupertipTitle"/>
                      <Description resid="TypicalReplyButton.SupertipText"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>onTypicalReplyButtonClicked</FunctionName>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
```

ここに、先ほど追加したボタンのテキストを使用した`Control`を追加します。

```
         <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="msgReadTabDefault">
                <Group id="msgReadCmdGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="msgReadTypicalReply">
                    <Label resid="TypicalReplyButton.Label"/>
                    <Supertip>
                      <Title resid="TypicalReplyButton.SupertipTitle"/>
                      <Description resid="TypicalReplyButton.SupertipText"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>onTypicalReplyButtonClicked</FunctionName>
                    </Action>
                  </Control>
                  <Control xsi:type="Button" id="newButton">
                    <Label resid="NewButton.Label"/>
                    <Supertip>
                      <Title resid="NewButton.SupertipTitle"/>
                      <Description resid="NewButton.SupertipText"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>onTypicalReplyButtonClicked</FunctionName>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
```

`<Control xsi:type="Button" id="newButton">`で指定した`id`は、この後設定ファイルで使用します。
`<Label resid="NewButton.Label"/>`の`resid`には、前の手順で作成したボタンのラベルのIDを指定します。
`<Supertip>`->`<Title resid="NewButton.SupertipTitle"/>`の`resid`には、前の手順で作成したボタンのスーパーチップのタイトルのIDを指定します。
`<Supertip>`->`<Description resid="NewButton.SupertipText"/>`の`resid`には、前の手順で作成したボタンのスーパーチップの説明のIDを指定します。
`<Action xsi:type="ExecuteFunction">` -> `<FunctionName>onTypicalReplyButtonClicked</FunctionName>`の値は`onTypicalReplyButtonClicked`とします。

次に、設定ファイル（`%APPDATA%\TypicalReply\TypicalReplyConfig.json`）を編集します。

現在の設定が以下のようになっているとします。

```
{
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "ButtonConfigList": [
                {
                    "Id": "msgReadTypicalReply",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        }
    ]
}
```

ButtonConfigListにButtonConfigを追加します。

`Id`に先程`<Control>`で追加した`newButton`を指定します。

```
{
    "Id": "newButton",
}
```

元のメッセージに対して返信するので、元の件名は残して、件名に対してリアクションのメッセージを追加します。
そのために、`Subject`は空にして元の件名が残るようにし、`SubjectPrefix`で件名の先頭にメッセージを追加します。

```
{
    "Id": "newButton",
    "SubjectPrefix": "[[最高！]]:"
}
```

同様に、元のメッセージに対して返信するので、元の本文は残して（引用状態にして）、本文にメッセージを追加します。
そのために、`Body`にメッセージを指定し、`QuoteType`に`true`を指定します。

```
{
    "Id": "newButton",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true
}
```

このボタンでは、送信者にのみ返信することとします。
そのために、`Recipients`に`["sender"]`を指定します。

```
{
    "Id": "newButton",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true,
    "Recipients": ["sender"]
}
```

また、送信先のドメインは自身が所属している`test.co.jp`のみに限定することとします。
そのために、`AllowedDomains`に`["all"]`を指定します。

```
{
    "Id": "newButton",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true,
    "Recipients": ["sender"],
    "AllowedDomains": ["test.co.jp"]
}
```

元のメッセージの添付は不要とします。
そのため、`ForwardType`は指定しません。

以上で作成した設定をButtonConfigListに追加します。

```
{
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "GalleryLabel": "定型返信",
            "ButtonConfigList": [
                {
                    "Id": "msgReadTypicalReply",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "newButton",
                    "SubjectPrefix": "[[最高！]]:",
                    "Body": "最高！",
                    "QuoteType": true,
                    "Recipients": ["sender"],
                    "AllowedDomains": ["test.co.jp"]
                }
            ]
        }
    ]
}
```

これで、定型返信のボタンの中に、「最高！」ボタンが追加されます。

!["「最高！」ボタン"](./documents/images/Awesome.PNG "「最高！」ボタン")

## 注意事項

### VSTO版との設定差異について

VSTO版に存在した以下の設定はOffice addin版には存在しません。

* TypiclReplyConfig: 設定のルート
  * Priority
    * 設定が複数存在する場合の優先度設定だが、現状では複数の設定を指定できないため

* Config: 各言語ごとの設定
  * GroupLabel
    * Officeアドインではグループのラベルは設定ファイルではなくマニフェストファイルで指定するため
  * TabMailInsertAfterMso
    * Officeアドインではアドインの挿入位置を調整できないため
  * TabReadInsertAfterMso
    * Officeアドインではアドインの挿入位置を調整できないため
  * ContextMenuInsertAfterMso
    * OfficeアドインではOutlookのコンテキストメニューが非サポートのため

* ButtonConfig: 定型返信ボタン設定。返信内容や返信先等の設定を行う。
  * Label
    * Officeアドインではボタンのラベルは設定ファイルではなくマニフェストファイルで指定するため
  * Size
    * Officeアドインではボタンのサイズを固定にすることができないため
  * Logo
    * 現在ボタンの画像の変更は非サポートのため