package com.tririga.custom;

import java.util.List;

public class GraphMessage {
    public String id;
    public String subject;
    public String receivedDateTime;
    public String sentDateTime;
    public GraphBody body;
    public GraphFrom from;
    public List<GraphToRecipient> getToRecipients;
    public boolean hasAttachments;
    public List<GraphEmailAttachment> attachments;
}
