package com.tririga.custom;

import java.util.List;

public class GraphMessage {
    public String id;
    public String subject;
    public String receivedDateTime;
    public String sentDateTime;
    public GraphBody body;
    public GraphPerson from;
    public List<GraphPerson> toRecipients;
    public boolean hasAttachments;
    public List<GraphEmailAttachment> attachments;

}
