/*
    Function to forward threads under a certain label
    This function has two inputs:
    - destinationemail: the email you want the labeled threads forwarded to
    - labeltoforward: the label you want to forward (function will include any sublabels as well)
    
    This function creates a new label that marks threads if they have been forwarded to
     the destination email by this funciton. This is intended to allow the user to re-run
     this function several times without re-forwarding emails that have already been sent.
    
    NOTE: It takes a few seconds (~2s) to send each email and so if there are hundreds emails,
      it could take several minutes to run this function. It would be best if you ran this function
      and did not use your computer for several minutes.
*/

function ForwardLabeledEmails() {

  //necessary variables (inputs)
  var destinationemail = "DESTINATION@EXAMPLE.COM"; //email to forward to
  var labeltoforward = "EXAMPLE LABEL"; //name of label to forward
  
  //Label for already forwarded threads
  var alftitle = "FWD " + destinationemail;
  var alreadyforwardedlabel = GmailApp.getUserLabelByName(alftitle);
  if ( alreadyforwardedlabel == null ) { //create label if doesn't exist
    alreadyforwardedlabel = GmailApp.createLabel(alftitle);
  } //if - alreadyforwardedlabel does not exist
  
  //message threads with the input label but not the already forwarded label
  var threads = GmailApp.search( _sublabelsearch( labeltoforward , alftitle ) );
  Logger.log(threads.length);
  for (var i = 0; i < threads.length ; i++) { 
    
    var t = threads[i]; //current thread in thread list
    
    //Obtain message information
    var msglist = t.getMessages(); //list of messages in thread
    var msg = msglist.reverse()[0]; //last message in thread (contains previous messages in body)
    var forwarded = msg.getBody(); //the body of the message we are forwarding.
    var oldsubject = msg.getSubject(); //Subject of message to forward (subject of last message in thread)
    var attachmentlist = _getFullAttachmentList( msglist ); //get all the attachments in the thread
    var head = _makeHeader( msg ); //construct header formatting in message to destination email
    
    //Forward message to (adds message to bottom of thread)
    // chose not to use this method in order to avoid cluttering up original threads
    // delete the /* and */ symbols to use this method (put them around the other method below)
    /*
    msg.forward( destinationemail , {
    replyTo: from , //sets replyto email to the original sender
    subject: oldsubject + "  [" + labeltoforward + "]", // customizes the subject
    htmlBody: header + forwarded, //forward message body
    attachments: attachmentlist
    });
    */
    
    //Send message in new thread (does not add message to bottom of original thread)
    MailApp.sendEmail(destinationemail,
                      oldsubject + "  [" + labeltoforward + "]",
                      "",
                      {
                      replyTo: msg.getFrom(),
                      attachments: attachmentlist,
                      htmlBody: head + forwarded
                      });
    
    //mark email as already forwarded
    t.addLabel( alreadyforwardedlabel );
    
  } //i - threads in label
}

/*
    Function to construct search string for label and sublables
     while excluding excludinglabel
    
    Inputs are both strings that match the name of a label already
     created
*/
function _sublabelsearch( baselabel , excludinglabel ) {
  
  //Get sublabels of input label (if they exist)
  // and construct search string
  var search = "";
  var sublabellist = GmailApp.getUserLabels().filter(
    function(label) {
      return label.getName().slice(0,baselabel.length) == baselabel;
    }
  ); //filters labels so that beginning of name is the same as input label
  for (var l = 0; l < sublabellist.length; l++) {
    var name = sublabellist[l].getName();
    search = search + "label:" + name.replace(/\s+/g,'-').replace(/\//g,'-').toLowerCase() + " ";
    if ( l < sublabellist.length-1 ) { //if not the last label in list
      search = search + "OR ";
    }
  }//l - label index
  search = search + "AND NOT label:" + excludinglabel;
  Logger.log(search); //check to make sure search is constructed correctly
  
  return search;
}

/*
    Function to obtain full attachment list from a list of message list
    
    Input is a list of GmailMessages
*/
function _getFullAttachmentList( messagelist ) {
  
  var attachlist = [];
  var att = ""; //temporary attachment variable
  for (var m=0; m < messagelist.length; m++) {
    att = messagelist[m].getAttachments(); //get attachments for message index m
    for (var a = 0; a < att.length; a++) {
      attachlist.push(att[a]); //add attachment index a for message index m to attachment list
    } //a - attaments in message
  } //m - messages in thread
  
  return attachlist;
}

/*
    Function to make the header for the forwarded message body
    
    Input is a GmailMessage that the information will be taken from
*/
function _makeHeader( msg ) {
  
  //get necessary information from message
  var date = msg.getDate(); //date and time of last message
  var from = msg.getFrom(); //sender of original email
  var to = msg.getTo(); //addressed to emails
  var ccd = msg.getCc(); //cc'd emails
  var bccd = msg.getBcc(); //bcc'd emails
  
  //construct header of forwarded email
  var header = "<div style='text-align: center;'>FORWARDED MESSAGE</div><hr><br>"; //Centered Title and horizontal line
  header = header + "From: " + from + "<br>To: " + to + "<br>"; //From and To lines
  if (ccd != "") {
    header = header + "Cc: " + ccd + "<br>"; //CC lines (if there are any)
  }
  if (bccd != "") {
    header = header + "Bcc: " + bccd + "<br>"; //BCC lines (if there are any)
  }
  header = header + "Date: " + Utilities.formatDate(date, "CDT", "EEE MMM d yyyy") + "<br>"; //Date line in Day Month Date Year format (e.g. Sat Apr 14 2018)
  
  return header;
}
