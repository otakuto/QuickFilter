global.buildAddOn = function(e)
{
  const accessToken = e.messageMetadata.accessToken
  GmailApp.setCurrentMessageAccessToken(accessToken)
  const messageId = e.messageMetadata.messageId
  const cards = []
  cards.push(mainCard(messageId))
  return cards
}

function mainCard(messageId)
{
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Create label and filter'))
  const section = CardService.newCardSection()

  const suggestionLabel = CardService.newSuggestions()

  const mail = GmailApp.getMessageById(messageId)

  let fromAddress = parseAddress(mail.getFrom())
  if (!fromAddress)
  {
    fromAddress = {
      name: mail.getFrom(),
      address: mail.getFrom()
    }
    suggestionLabel.addSuggestion(mail.getFrom())
  }
  else
  {
    suggestionLabel.addSuggestion(fromAddress.name).addSuggestion(fromAddress.address)
  }

  const labels = GmailApp.getUserLabels()
  for (const e of labels)
  {
    suggestionLabel.addSuggestion(e.getName())
  }

  const textInputLabelName = CardService.newTextInput()
    .setTitle('Label Name')
    .setFieldName('labelName')
    .setValue(fromAddress.name)
    .setSuggestions(suggestionLabel)
  section.addWidget(textInputLabelName)

  const dropDownLabelListVisibility = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Show in label list')
    .setFieldName('labelListVisibility')
    .addItem('Show', 'labelShow', false)
    .addItem('Show if unread', 'labelShowIfUnread', true)
    .addItem('Hide', 'labelHide', false)
  section.addWidget(dropDownLabelListVisibility)

  const textInputFromAddress = CardService.newTextInput()
    .setTitle('From')
    .setFieldName('fromAddress')
    .setValue(fromAddress.address)
  section.addWidget(textInputFromAddress)

  const buttonMatchAddresses = CardService.newTextButton()
    .setText('View matching addresses')
    .setOnClickAction(CardService.newAction().setFunctionName('onClickActionButtonMatchAddresses'))
  section.addWidget(buttonMatchAddresses)

  const checkboxUnmarkInbox = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('unmarkInbox')
    .addItem('Skip the Inbox (Archive it)', true, false)
  section.addWidget(checkboxUnmarkInbox)

  const checkboxUnmarkUnread = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('unmarkUnread')
    .addItem('Mark as read', true, false)
  section.addWidget(checkboxUnmarkUnread)

  const checkboxMarkStar = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('markStar')
    .addItem('Star it', true, false)
  section.addWidget(checkboxMarkStar)

  const checkboxMarkTrash = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('markTrash')
    .addItem('Deconste it', true, false)
  section.addWidget(checkboxMarkTrash)

  const checkboxUnmarkSpam = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('unmarkSpam')
    .addItem('Never send it to Spam', true, false)
  section.addWidget(checkboxUnmarkSpam)

  const checkboxMarkImportant = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('markImportant')
    .addItem('Always mark it as important', true, false)
  section.addWidget(checkboxMarkImportant)

  const checkboxUnmarkImportant = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('unmarkImportant')
    .addItem('Never mark it as important', true, false)
  section.addWidget(checkboxUnmarkImportant)

  const checkboxApplyFilter = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('applyFilter')
    .addItem('Also apply filter to matching conversations', true, false)
  section.addWidget(checkboxApplyFilter)

  const buttonCreate = CardService.newTextButton()
    .setText('create')
    .setOnClickAction(CardService.newAction().setFunctionName('onClickActionButtonCreate'))

  const buttonPreview = CardService.newTextButton()
    .setText('preview')
    .setOnClickAction(CardService.newAction().setFunctionName('onClickActionButtonPreview'))

  section.addWidget(CardService.newButtonSet().addButton(buttonPreview).addButton(buttonCreate))

  card.addSection(section)
  return card.build()
}

global.onClickActionButtonCreate = function(e)
{
  if (e.formInput.markImportant && e.formInput.unmarkImportant)
  {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setType(CardService.NotificationType.ERROR)
        .setText('Marking of Important is invalid.'))
      .build()
  }

  const labelName = e.formInput.labelName

  const label = Gmail.newLabel()
  label.name = labelName
  label.labelListVisibility = e.formInput.labelListVisibility
  label.messageListVisibility = 'show'
  try
  {
    Gmail.Users.Labels.create(label, 'me')
  }
  catch(e)
  {
  }

  const addLabelIds = labelNamesToIDs([labelName])

  if (e.formInput.markStar)
  {
    addLabelIds.push('STARRED')
  }

  if (e.formInput.markTrash)
  {
    addLabelIds.push('TRASH')
  }

  if (e.formInput.markImportant)
  {
    addLabelIds.push('IMPORTANT')
  }

  const removeLabelIds = []
  if (e.formInput.unmarkInbox)
  {
    removeLabelIds.push('INBOX')
  }

  if (e.formInput.unmarkUnread)
  {
    removeLabelIds.push('UNREAD')
  }

  if (e.formInput.unmarkSpam)
  {
    removeLabelIds.push('SPAM')
  }

  if (e.formInput.unmarkImportant)
  {
    removeLabelIds.push('IMPORTANT')
  }

  const from = e.formInput.fromAddress
  try
  {
    createFilter(from, addLabelIds, removeLabelIds)
  }
  catch(e)
  {
  }

  if (e.formInput.applyFilter)
  {
    const q = 'from:(' + from + ')'
    applyFilterInbox(q, addLabelIds, removeLabelIds)
    GmailApp.getInboxThreads(0, 1)[0].refresh()
  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setType(CardService.NotificationType.INFO)
      .setText('Successful.'))
    .build()
}

function parseAddress(address)
{
  const regex = /(.+) \<([^\@]+\@[^\>]+)\>/
  const match = regex.exec(address)
  if (match)
  {
    return {
      name: match[1].replace(/\"/g, ''),
      address: match[2]
    }
  }
  else
  {
    return null
  }
}

global.onClickActionButtonMatchAddresses = function(e)
{
  const fromAddress = e.formInput.fromAddress

  const addresses = matchAddressesInbox('from:(' + fromAddress + ')')

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle((addresses.length > 0) ? (addresses.length + ' matching addresses.') : ('No addresses matched your search.')))

  let text = ''
  for (const address of addresses)
  {
    //const count = counts[address]
    //const a = parseAddress(address)
    //if (a)
    //{
    //  text += '<font color=\"#9E9E9E\">' + a.name + '</font><br>' + a.address + '<br>'
    //}
    //else
    //{
    //  text += address + '<br>'
    //}
    //text += '<font color=\"#1257e0\">' + counts[address] + ' messages.</font><br><br>'

    text += address + '<br>'
  }
  const textParagraph = CardService.newTextParagraph().setText(text)

  card.addSection(CardService.newCardSection().addWidget(textParagraph)).build()

  const nav = CardService.newNavigation().pushCard(card.build())
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build()
}

global.onClickActionButtonPreview = function(e)
{
  const fromAddress = e.formInput.fromAddress

  const count = countMessagesInbox('from:(' + fromAddress + ')')
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(count ? (count + ' matching messages.') : ('No messages matched your search.')))

  const threads = GmailApp.search('from:(' + fromAddress + ')', 0, 20)
  for (const t of threads)
  {
    const m = t.getMessages()[0]
    const header = CardService.newKeyValue().setTopLabel(m.getFrom().replace(/\"/g, '')).setContent(m.getSubject())
    const subject = CardService.newKeyValue().setMultiline(true).setTopLabel('Subject:').setContent(m.getSubject())

    const from = CardService.newKeyValue().setMultiline(true).setTopLabel('From:')
    const fromAddress = parseAddress(m.getFrom())
    if (fromAddress)
    {
      from.setContent(fromAddress.address)
    }
    else
    {
      from.setContent(m.getFrom())
    }

    const to = CardService.newKeyValue().setMultiline(true).setTopLabel('To:')
    const toAddress = parseAddress(m.getTo())
    if (toAddress)
    {
      to.setContent(toAddress.address)
    }
    else
    {
      to.setContent(m.getTo())
    }

    const openMail = CardService.newTextButton()
      .setText('Open')
      .setOpenLink(CardService.newOpenLink()
        .setUrl(t.getPermalink())
        .setOpenAs(CardService.OpenAs.OVERLAY))

    const section = CardService.newCardSection()
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(1)
      .addWidget(header)
      .addWidget(from)
      .addWidget(to)
      .addWidget(subject)
      .addWidget(openMail)
    card.addSection(section)
  }

  const nav = CardService.newNavigation().pushCard(card.build())
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build()
}

function createFilter(from, addLabelIds, removeLabelIds)
{
  const filter = Gmail.newFilter()

  filter.criteria = Gmail.newFilterCriteria()
  filter.criteria.from = from

  filter.action = Gmail.newFilterAction()
  filter.action.addLabelIds = addLabelIds
  filter.action.removeLabelIds = removeLabelIds

  Gmail.Users.Settings.Filters.create(filter, 'me')
}

function countMessagesInbox(q)
{
  let s = 0
  const f = function(messages)
  {
    s += messages.length
  }

  forEachInboxMessages(q, f)

  return s
}

function matchAddressesInbox(q)
{
  const addresses = []
  while (true)
  {
    const exclude = (addresses.length > 0) ? ' -{from:(' + addresses.reduce((x, y) => x + '|' + y)+ ')}' : ''
    const m = Gmail.Users.Messages.list('me', {q: q + exclude, maxResults: 1})
    if (m.messages && m.messages.length > 0)
    {
      const metadata = Gmail.Users.Messages.get('me', m.messages[0].id, {format: 'metadata'})
      for (const e of metadata.payload.headers)
      {
        if (e.name === 'From')
        {
          const address = parseAddress(e.value)
          if (address)
          {
            addresses.push(address.address)
          }
          else
          {
            addresses.push(e.value)
          }
        }
      }
    }
    else
    {
      return addresses
    }
  }
}

function labelNamesToIDs(names)
{
  const label = {}
  Gmail.Users.Labels.list('me').labels.forEach(function(e)
  {
    label[e.name] = e.id
  })

  const ids = []
  names.forEach(function(e)
  {
    if (label[e])
    {
      ids.push(label[e])
    }
    else
    {
      throw new Error('"' + e + '" label is not available.')
    }
  })

  return ids
}

function applyFilterInbox(q, addLabelIds, removeLabelIds)
{
  const f = function(messages)
  {
    const batch = Gmail.newBatchModifyMessagesRequest()
    batch.ids = messages.map(function(e){return e.id})
    batch.addLabelIds = addLabelIds
    batch.removeLabelIds = removeLabelIds
    Gmail.Users.Messages.batchModify(batch, 'me')
  }

  forEachInboxMessages(q, f)
}

function forEachInboxThreads(q, f)
{
  const max = 200
  let start = 0
  while (true)
  {
    const threadList = GmailApp.search(q, start, max)
    if (threadList && threadList.length > 0)
    {
      f(threadList)
      start += max
    }
    else
    {
      break
    }
  }
}

function forEachInboxMessages(q, f)
{
  let pageToken
  do
  {
    const messagesList = Gmail.Users.Messages.list('me', {q: q, pageToken: pageToken})
    if (messagesList.messages && messagesList.messages.length > 0)
    {
      f(messagesList.messages)
    }
    pageToken = messagesList.nextPageToken
  }
  while (pageToken)
}

