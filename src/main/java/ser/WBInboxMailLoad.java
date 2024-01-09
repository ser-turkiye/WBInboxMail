package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.collections4.IteratorUtils;
import org.json.JSONObject;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

import static ser.Utils.loadTableRows;
import static ser.Utils.saveWBInboxExcel;


public class WBInboxMailLoad extends UnifiedAgent {

    ISession ses;
    IDocumentServer srv;
    IBpmService bpm;

    JSONObject mailTemplates = new JSONObject();
    JSONObject projects = new JSONObject();
    ProcessHelper helper;

    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        (new File(Conf.WBInboxMailPaths.MainPath)).mkdir();
        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        bpm = getBpm();
        ses = getSes();
        srv = ses.getDocumentServer();

        try {
            helper = new ProcessHelper(ses);
            JSONObject mcfg = Utils.getMailConfig(ses, srv, "");

            List<IWorkbasket> wbs = bpm.getWorkbaskets();
            for (IWorkbasket wb : wbs){
                IWorkbasket swb = bpm.getWorkbasket(wb.getID());
                runWorkbasket(swb, mcfg);

            }

            System.out.println("Tested.");
        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        System.out.println("Finished");
        return resultSuccess("Ended successfully");
    }
    private IInformationObject getProject(String prjn){
        if(projects.has(prjn)){return (IInformationObject) projects.get(prjn);}
        if(projects.has("!" + prjn)){return null;}

        IInformationObject iprj = Utils.getProjectWorkspace(prjn, helper);
        if(iprj == null){
            projects.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        projects.put(prjn, iprj);
        return iprj;
    }
    private IDocument getMailTplDocument(String prjn) throws Exception{
        if(mailTemplates.has(prjn)){return (IDocument) mailTemplates.get(prjn);}
        if(mailTemplates.has("!" + prjn)){return null;}

        IInformationObject prjt = getProject(prjn);
        if(prjt == null){return null;}

        IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.Project, ses, srv);
        if(dtpl == null){
            mailTemplates.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        mailTemplates.put(prjn, dtpl);
        return dtpl;
    }
    private void runWorkbasket(IWorkbasket swb, JSONObject mcfg) throws Exception {
        String wbMail = swb.getNotifyEMail();

        System.out.println("WB : " + swb.getName());
        System.out.println(" *** mail : " + wbMail);
        System.out.println(" *** fullname : " + swb.getFullName());
        System.out.println(" *** accessible : " + swb.isAccessible());

        if(!swb.isAccessible()){return;}

        JSONObject prjDocs = new JSONObject();
        if(wbMail == null || wbMail.isEmpty()){return;}

        IWorkbasketContent wbco = swb.getWorkbasketContent();
        List<ITask> tasks = wbco.getTasks();
        if(tasks.size() < 1){return;}


        System.out.println("    -> start ");
        int tcnt = 0;
        JSONObject npjs = new JSONObject();
        for(ITask task : tasks){

            tcnt++;
            System.out.println(" *** task [" + tcnt + "] : " + task.getDisplayName());

            String clid = task.getClassID();
            System.out.println(" *** clid [" + tcnt + "] : " + clid);
            if(!clid.equals(Conf.ClassIDs.Transmittal)
                    && !clid.equals(Conf.ClassIDs.SubReview)
                    && !clid.equals(Conf.ClassIDs.ReviewMain)){continue;}

            IProcessInstance proi = task.getProcessInstance();
            if(proi == null){continue;}
            System.out.println(" *** proi [" + tcnt + "] : " + proi.getDisplayName());

            IDocument wdoc = (IDocument) proi.getMainInformationObject();
            if(wdoc == null){continue;}
            System.out.println(" *** wdoc [" + tcnt + "] : " + wdoc.getDisplayName());

            String prjn = wdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn == null || prjn.isEmpty()){continue;}
            System.out.println(" *** prjn [" + tcnt + "] : " + prjn);

            if(npjs.has(prjn)){continue;}
            if(getMailTplDocument(prjn) == null){
                npjs.put(prjn, prjn);
                continue;
            }

            if(!prjDocs.has(prjn)){
                prjDocs.put(prjn, new JSONObject());
            }
            JSONObject docs = (JSONObject) prjDocs.get(prjn);

            if(docs.has(task.getID())){continue;}

            System.out.println("    -> class-name : " + task.getName());
            System.out.println("    -> class-id.doc : " + wdoc.getClassID());
            System.out.println("    -> class-id.task : " + clid);
            System.out.println("    -> display : " + wdoc.getDisplayName());

            docs.put(task.getID(), task);
        }


        if(prjDocs.length() < 1){return;}
        List<String> prjs = IteratorUtils.toList(prjDocs.keys());
        for(String prjn : prjs){
            if(!prjDocs.has(prjn)){continue;}
            JSONObject docs = (JSONObject) prjDocs.get(prjn);
            if(docs.length() < 1){continue;}


            String uniqueId = UUID.randomUUID().toString();
            IDocument dtpl = getMailTplDocument(prjn);
            if(dtpl == null){continue;}

            String mailExcelPath = Utils.exportDocument(dtpl, Conf.WBInboxMailPaths.MainPath, Conf.MailTemplates.Project + "@" + prjn + "[" + uniqueId + "]");
            List<String> dids = IteratorUtils.toList(docs.keys());

            loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, dids.size());

            JSONObject mbms = new JSONObject();
            mbms.put("Fullname", swb.getFullName());
            mbms.put("Count", dids.size() + "");
            int dcnt = 0;
            for(String zdid : dids){
                dcnt++;
                if(!docs.has(zdid)){continue;}
                ITask xtsk = (ITask) docs.get(zdid);
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                if(orwb != null && orwb.getID() != swb.getID()){
                    owbn = orwb.getFullName();
                }


                Date tbgn = null, tend = new Date();
                if(xtsk.getReadyDate() != null){
                    tbgn = xtsk.getReadyDate();
                }

                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }

                String rcvf = "", rcvo = "";
                if(xtsk.getPreviousWorkbasket() != null){
                    rcvf = xtsk.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                String mdno = "", mdrn = "", mdnm = "";
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                    mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                    mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                    mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }

                mbms.put("Task" + dcnt, xtsk.getName());
                mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                mbms.put("DurDay" + dcnt, durd + "");
                mbms.put("DurHour" + dcnt, durh + "");
                mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
            }

            saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                    Conf.WBInboxMailPaths.MainPath + "/" + Conf.MailTemplates.Project + "@" + prjn + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            mail.put("Subject",
                "Workbox Notification. For {ProjectNo},  ({Count}) Task is waiting your action"
                        .replace("{ProjectNo}", prjn)
                        .replace("{Count}", dids.size() + "")
            );
            //mail.put("Subject", "Reminder > " + prjn + " / " + swb.getFullName());
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                Utils.sendHTMLMail(ses, srv, mcfg, mail);
            }catch(Exception ex){
                System.out.println("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }


        System.out.println("    -> finish ");
    }

}