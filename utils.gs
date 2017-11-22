var monthDitcionnary = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
var dayDitcionnary   = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];

function getHumanReadableDate(dateString) 
{
  var date = new Date(dateString);
  var humanReadableDateString = dayDitcionnary[date.getDay()] + " " + date.getDate() + " " + monthDitcionnary[date.getMonth()] + " " + date.getYear();
  
  return humanReadableDateString;
}

var bccMails = 'bibien22@hotmail.com'
var replyToAddr = "sonneur.roc14@gmail.com"
var refLinkCal = 'https://docs.google.com/spreadsheets/d/1LuxJrXUaHV6ncIQ0h6R2aDvyMQuHKEe-wVS7ko4mnOU/edit'
var addressMailReferents           = "referents@roc14.org";
var addressMailReferentsNocturnes  = "referents.nocturnes@roc14.org";
var addressMailAnnulation          = "bureau@roc14.org";
var noon    = 0;
var normal  = 1;
var night   = 2;
var weekend = 3;

function createFirstCallMessage(date, timeSlot)
{
  if(timeSlot == noon)
  {
    var strTimeSlot = " midi ";
    var mailAddr = addressMailReferents;
  }
  else if(timeSlot == normal)
  {
    var strTimeSlot = " soir ";
    var mailAddr = addressMailReferents;
  }  
  else if(timeSlot == night)
  {
    var strTimeSlot = " soir (nocturne) ";
    var mailAddr = addressMailReferentsNocturnes;
  }
  else
  {
    var strTimeSlot = " ";
    var mailAddr = addressMailReferents;
  }

  return Message = 
  {
    name : 'Sonneur référents',
    subject : '[referents] pas de référents',
    to : mailAddr,
    bcc: bccMails,
    replyTo : replyToAddr,
    //htmlBody required for OVH stupid mailing list (or at least it works...). Extracted from a google message.
    htmlBody: ' <div><span style="font-size:12.8px">Bonjour,</span>' +
    '<br><br>' +
    '<span style="font-size:12.8px">Il n&#39;y a pas de référent d&#39;inscrit pour demain' + strTimeSlot
    + '(' + date + ').' + 
    '</span><br><br>' +
    '<span style="font-size:12.8px">Je rappelle le lien vers le classeur à remplir :</span>' +
    '<br><a href="' + refLinkCal + '" rel="noreferrer" target="_blank" style="font-size:12.8px">' + refLinkCal + '</a>' +
    '<br><br>' +
    '<span style="font-size:12.8px">L&#39;équipe des référents.</span><br>' +
    '</div>'
  };
}

function sendFirstCallMail(completeDateStr, timeSlot)
{
  MailApp.sendEmail(createFirstCallMessage(getHumanReadableDate(completeDateStr), timeSlot));
}

function createCancellationMessage(date, timeSlot)
{
  if(timeSlot == noon)
  {
    var strTimeSlot = " séance de demain midi ";
  }
  else if(timeSlot == normal)
  {
    var strTimeSlot = " séance de ce soir ";
  }  
  else if(timeSlot == night)
  {
    var strTimeSlot = " nocturne de ce soir ";
  }
  else if(timeSlot == weekend)
  {
    var strTimeSlot = " séance de demain ";
  }
  else
  {
    var strTimeSlot = " ";
  }

  return Message = 
  {
    name : 'Sonneur référents',
    subject : '[referents] pas de référents',
    to : addressMailAnnulation,
    bcc: bccMails,
    replyTo : replyToAddr,
    //htmlBody required for OVH stupid mailing list (or at least it works...). Extracted from a google message.
    htmlBody: ' <div><span style="font-size:12.8px">Bonjour,</span>' +
    '<br><br>' +
    '<span style="font-size:12.8px">Il n&#39;y a toujours pas de référent d&#39;inscrit pour la' + strTimeSlot
    + '(' + date + ').' + 
    '</span><br><br>' +
    '<span style="font-size:12.8px">L&#39;annulation du créneau s&#39;impose ?</span>' +
    '</span><br><br>' +
    '<br><a href="' + refLinkCal + '" rel="noreferrer" target="_blank" style="font-size:12.8px">' + refLinkCal + '</a>' +
    '<br><br>' +
    '<span style="font-size:12.8px">Le sonneur.</span><br>' +
    '</div>'
  };
}

function sendCancellationMail(completeDateStr, timeSlot)
{
  MailApp.sendEmail(createCancellationMessage(getHumanReadableDate(completeDateStr), timeSlot));
}

