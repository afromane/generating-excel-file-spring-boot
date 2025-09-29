package com.afromane.exportToExcel.model;

public class PeopleReviewData {
    private String postesOccupants;
    private String datePriseFonction;
    private String anciennete;
    private String potentielEvolution;
    private String performance;
    private String statut;
    private String revalorisation;
    private String promotion;
    private String primeExceptionnelle;
    private String autresAvantages;
    private String formations;
    private String actionSuivi2024;
    private String commentaires;

    // Default constructor
    public PeopleReviewData() {}

    // Full constructor
    public PeopleReviewData(String postesOccupants, String datePriseFonction, String anciennete,
                            String potentielEvolution, String performance, String statut,
                            String revalorisation, String promotion, String primeExceptionnelle,
                            String autresAvantages, String formations, String actionSuivi2024,
                            String commentaires) {
        this.postesOccupants = postesOccupants;
        this.datePriseFonction = datePriseFonction;
        this.anciennete = anciennete;
        this.potentielEvolution = potentielEvolution;
        this.performance = performance;
        this.statut = statut;
        this.revalorisation = revalorisation;
        this.promotion = promotion;
        this.primeExceptionnelle = primeExceptionnelle;
        this.autresAvantages = autresAvantages;
        this.formations = formations;
        this.actionSuivi2024 = actionSuivi2024;
        this.commentaires = commentaires;
    }

    // Getters and Setters
    public String getPostesOccupants() { return postesOccupants; }
    public void setPostesOccupants(String postesOccupants) { this.postesOccupants = postesOccupants; }

    public String getDatePriseFonction() { return datePriseFonction; }
    public void setDatePriseFonction(String datePriseFonction) { this.datePriseFonction = datePriseFonction; }

    public String getAnciennete() { return anciennete; }
    public void setAnciennete(String anciennete) { this.anciennete = anciennete; }

    public String getPotentielEvolution() { return potentielEvolution; }
    public void setPotentielEvolution(String potentielEvolution) { this.potentielEvolution = potentielEvolution; }

    public String getPerformance() { return performance; }
    public void setPerformance(String performance) { this.performance = performance; }

    public String getStatut() { return statut; }
    public void setStatut(String statut) { this.statut = statut; }

    public String getRevalorisation() { return revalorisation; }
    public void setRevalorisation(String revalorisation) { this.revalorisation = revalorisation; }

    public String getPromotion() { return promotion; }
    public void setPromotion(String promotion) { this.promotion = promotion; }

    public String getPrimeExceptionnelle() { return primeExceptionnelle; }
    public void setPrimeExceptionnelle(String primeExceptionnelle) { this.primeExceptionnelle = primeExceptionnelle; }

    public String getAutresAvantages() { return autresAvantages; }
    public void setAutresAvantages(String autresAvantages) { this.autresAvantages = autresAvantages; }

    public String getFormations() { return formations; }
    public void setFormations(String formations) { this.formations = formations; }

    public String getActionSuivi2024() { return actionSuivi2024; }
    public void setActionSuivi2024(String actionSuivi2024) { this.actionSuivi2024 = actionSuivi2024; }

    public String getCommentaires() { return commentaires; }
    public void setCommentaires(String commentaires) { this.commentaires = commentaires; }
}