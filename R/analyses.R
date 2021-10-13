library(ggplot2)
library(pander)
library(parallel)
library(readxl)
library(writexl)

options(mc.cores = detectCores() - 1) 

# Working directory
setwd("~/Projects/Consultations/Palmero David (NEC)")

# Output directory
outdir <- paste0("results/analyses_", format(Sys.Date(), "%Y%m%d"))
if (!dir.exists(outdir)) dir.create(outdir)

# ------------------- Data importation and preprocessing -------------------- #

# Import data
fileName <- "data-raw/Données NEC 06.10.2021.xlsx"
dta <- as.data.frame(read_xlsx("data-raw/Données NEC 06.10.2021.xlsx"))

# Rename variables
names(dta) <- sub("\\s*\\([^\\)]+\\)", "", names(dta))
names(dta) <- iconv(names(dta), to = "ASCII//TRANSLIT")
names(dta) <- gsub(" ", ".", sub("D'", "", names(dta)))

# Recoding
dta$Groupe <- factor(dta$Groupe, 1:3)
dta$NEC <- dta$NEC - 1
dta$Type.accouchement <- factor(dta$Type.accouchement, 0:1,
                                c("VoieBasse", "Cesarienne"))
dta$Sexe <- factor(dta$Sexe, 1:2, c("F", "H"))
dta$Surfactant <- dta$Surfactant - 1
dta$Corticoides.antenatal <- factor(dta$Corticoides.antenatal, c(0, 2:3),
                                    c("Non", "Partiel", "Complet"))
dta$HIV <- factor(dta$HIV, 0:4)
dta$CA <- dta$CA - 1
dta$LP <- factor(dta$LP, 0:2, c("Non", "Echogenique", "Cystique"))
dta$Age.corrige.RD.sans.VE <-
  ifelse(dta$Age.corrige.RD <= 100, dta$Age.corrige.RD, NA)
dta$Log.Age.corrige.RD <- log(dta$Age.corrige.RD)
dta$BoxCox.Age.corrige.RD <- 1 - 1 / dta$Age.corrige.RD


# --------------------------- Desciptive analyses --------------------------- #

# Type of variables
V_cat <- names(dta)[sapply(dta, class) == "factor"]
V_bin <- names(dta)[sapply(dta, function(x) all(x %in% 0:1 | is.na(x)))]
V_num <- names(dta)[sapply(dta, class) == "numeric" & 
                      !(names(dta) %in% c("Groupe", V_bin))]
V_num <- c(V_num[1], V_num[(length(V_num) - 2):length(V_num)],
           V_num[2:(length(V_num) - 3)])

# Descriptive analyses - Numeric variables
rows <- mclapply(setNames(V_num, V_num), function(v) {
  x <- list(
    Tous = dta[[v]],
    Groupe1 = dta[dta$Groupe == 1, v],
    Groupe2 = dta[dta$Groupe == 2, v],
    Groupe3 = dta[dta$Groupe == 3, v]
  )
  x <- lapply(x, function(z) z[!is.na(z)])
  n <- length
  q25 <- function(z) unname(quantile(z, 0.25))
  q75 <- function(z) unname(quantile(z, 0.75))
  fct <- c("n", "mean", "sd", "min", "q25", "median", "q75", "max")
  r <- do.call(c, lapply(names(x), function(u) {
    r <- sapply(fct, function(f) {
      if (length(x[[u]]) > 0) get(f)(x[[u]]) else NA
    })
    names(r) <- paste(u, names(r), sep = ".")
    return(r)
  }))
  lg <- do.call(rbind, lapply(names(x), function(u) {
    if (length(x[[u]]) > 0) data.frame(grp = u, y = x[[u]]) else NULL
  }))
  lg$grp <- droplevels(factor(
    lg$grp, c("Tous", "Groupe1", "Groupe2", "Groupe3")))
  fig <- ggplot(lg, aes(x = grp, y = y)) +
    geom_boxplot() +
    labs(title = v, x = "", y = v)
  r <- cbind(
    data.frame(Variable = v),
    t(r),
    G1_G2_t.test_p.value =
      tryCatch(t.test(x$Groupe1, x$Groupe2)$p.value, error = function(err) NA),
    G1_G2_wilcox_p.value =
      tryCatch(wilcox.test(x$Groupe1, x$Groupe2, exact = FALSE)$p.value,
               error = function(err) NA),
    G1_G3_t.test_p.value =
      tryCatch(t.test(x$Groupe1, x$Groupe3)$p.value, error = function(err) NA),
    G1_G3_wilcox_p.value =
      tryCatch(wilcox.test(x$Groupe1, x$Groupe3, exact = FALSE)$p.value,
               error = function(err) NA),
    G2_G3_t.test_p.value =
      tryCatch(t.test(x$Groupe2, x$Groupe3)$p.value, error = function(err) NA),
    G2_G3_wilcox_p.value =
      tryCatch(wilcox.test(x$Groupe2, x$Groupe3, exact = FALSE)$p.value,
               error = function(err) NA)
  )
  attr(r, "fig") <- fig
  return(r)
})
tbl_num <- do.call(rbind, rows)
figs_num <- lapply(rows, function(r) attr(r, "fig"))
rm(rows)

# Descriptive analyses - Categorical and binary variables
V <- c(V_bin, V_cat)
tbl_cat <- do.call(rbind, mclapply(setNames(V, V), function(v) {
  x <- list(
    Tous = dta[[v]],
    Groupe1 = dta[dta$Groupe == 1, v],
    Groupe2 = dta[dta$Groupe == 2, v],
    Groupe3 = dta[dta$Groupe == 3, v]
  )
  Merge <- function(x1, x2) merge(x1, x2, by = "value", all = TRUE,
                                  sort = FALSE)
  r <- Reduce(Merge, lapply(names(x), function(u) {
    r <- table(x[[u]])
    r <- cbind(n = r, prop = prop.table(r))
    colnames(r) <- paste(u, colnames(r), sep = ".")
    cbind(data.frame(value = rownames(r)), r)
  }))
  tbl1 <- rbind(data.frame(g = 1, v = x$Groupe1),
                data.frame(g = 2, v = x$Groupe2))
  tbl1 <- table(tbl1$g, tbl1$v)
  tbl2 <- rbind(data.frame(g = 1, v = x$Groupe1),
                data.frame(g = 2, v = x$Groupe3))
  tbl2 <- table(tbl2$g, tbl2$v)
  tbl3 <- rbind(data.frame(g = 1, v = x$Groupe2),
                data.frame(g = 2, v = x$Groupe3))
  tbl3 <- table(tbl3$g, tbl3$v)
  na <- rep(NA, nrow(r) - 1)
  cbind(
    variable = c(v, na),
    r,
    G1_G2_chisq.test_p.value = c(chisq.test(tbl1)$p.value, na),
    G1_G2_fisher.test_p.value = c(fisher.test(tbl1)$p.value, na),
    G1_G3_chisq.test_p.value = c(chisq.test(tbl2)$p.value, na),
    G1_G3_fisher.test_p.value = c(fisher.test(tbl2)$p.value, na),
    G2_G3_chisq.test_p.value = c(chisq.test(tbl3)$p.value, na),
    G2_G3_fisher.test_p.value = c(fisher.test(tbl3)$p.value, na)
  )
}))
rm(V)

# Export descriptive analyses
write_xlsx(list(numeric = tbl_num, categorical = tbl_cat),
           file.path(outdir, "analyses_descriptives.xlsx"))
f1 <- file.path(outdir, "analyses_descriptives.pdf")
pdf(f1)
for (fig in figs_num) print(fig)
dev.off()
f2 <- "/tmp/tmp_bookmarks.txt"
for(i in 1:length(figs_num)) {
  write("BookmarkBegin", file = f2, append = (i!=1))
  s <- paste("BookmarkTitle:", names(figs_num)[i])
  write(s, file = f2, append = TRUE)
  write("BookmarkLevel: 1", file = f2, append = TRUE)
  write(paste("BookmarkPageNumber:", i), file = f2, append = TRUE)
}
system(paste("pdftk", f1, "update_info", f2, "output /tmp/tmp_fig.pdf"))
system(paste("mv /tmp/tmp_fig.pdf", f1, "&& rm", f2))
rm(fig, f1, f2, i, s)

# ------------------------------- Regressions ------------------------------- #

# Regressions: Dependent and independent variables
Y <- c("NEC", "Late.onset.sepsis", "Mortalite", "Age.corrige.RD",
       "Age.corrige.RD.sans.VE", "Log.Age.corrige.RD", "BoxCox.Age.corrige.RD")
X <- names(dta)[!(names(dta) %in% Y)]

# Univariable regressions
uvreg <- mclapply(setNames(Y, Y), function(y) {
  do.call(rbind, lapply(X, function(x) {
    fml <- as.formula(paste(y, "~", x))
    fam <- if (grepl("Age\\.corrige", y)) "gaussian" else "binomial"
    fit <- do.call("glm", list(formula = fml, family = fam, data = quote(dta)))
    b <- cbind(coef(fit), suppressMessages(confint(fit)))
    if (fam == "binomial") {
      b <- exp(b)
      colnames(b)[1] <- "odds ratio"
    } else {
      colnames(b)[1] <- "beta"
    }
    pv <- coef(summary(fit))[, 4]
    s <- ifelse(pv > 0.1, NA, ifelse(pv > 0.05, "≤0.1",
           ifelse(pv > 0.01, "≤0.05", ifelse(pv > 0.001, "≤0.01" ,"≤0.001"))))
    s[1] <- NA
    b <- cbind(b, `p-value` = pv, significativity = s)
    na <- rep(NA, nrow(b) - 1)
    r <- data.frame(
      `dependent variable` = c(y, na),
      `independent variable` = c(x, na),
      `number of observations` = c(nrow(fit$model), na),
      coefficient = rownames(b)
    )
    cbind(r, b)
  }))
})
write_xlsx(uvreg, file.path(outdir, "regressions_univariables.xlsx"))

# Univariable regressions - diagnostic plots
pdf(file.path(outdir, "regressions_univariables.pdf"))
for (y in Y[grep("Age", Y)]) {
  for (x in names(dta)[!(names(dta) %in% Y)]) {
    fml <- paste(y, "~", x)
    fit <- lm(as.formula(fml), dta)
    par(mfrow = c(2, 2))
    for (i in 1:4) plot(fit, i)
    par(mfrow = c(1, 1))
    mtext(fml, outer = TRUE, line = -1.8, cex = 1)
  }
}
dev.off()
rm(i, fit, fml, x, y)

# Regressions with two dependent variables (Groupe + other)
mvreg <- mclapply(1:4, function(k) {
  mclapply(list(Tous = 1:3, G12 = 1:2, G23 = 2:3), function(g) {
    sdta <- subset(dta, Groupe %in% g)
    sdta$Groupe <- droplevels(sdta$Groupe)
    mclapply(setNames(Y, Y), function(y) {
      X <- X[X != "Groupe"]
      if (all(2:3 %in% g)) {
        X <- X[X != "Annee.hospitalisation"]
      }
      if (k == 1) {
        X <- c(NA, X)
      } else if (k == 2) {
        X <- combn(X, 2, simplify = FALSE)
      } else {
        X <- list(X)
      }
      do.call(rbind, lapply(X, function(x) {
        if (all(is.na(x))) {
          x_fml <- "Groupe"
          sdta <- na.omit(sdta[c(y, "Groupe")])
        } else {
          x_fml <- paste(c("Groupe", x), collapse = " + ")
          sdta <- na.omit(sdta[c(y, "Groupe", x)])
        }
        #print(paste("k =", k, "/ g =", paste(g, collapse = ","), "/ fml =",
        #            y, "~", x_fml))
        fml <- as.formula(paste(y, "~", x_fml))
        fam <- if (grepl("Age\\.corrige", y)) "gaussian" else "binomial"
        z <- paste('do.call("glm", list(formula = fml, family = fam,',
                   'data = quote(sdta)))')
        if (k == 4) {
          z <- paste0('step(', z, ', scope = list(lower = "NEC ~ Groupe"),',
                      ' trace = FALSE)')
        }
        z <- evals(z, env = environment(), graph.dir = "/tmp/R_pander")[[1]]
        fit <- z$result
        if (k == 4) x_fml <- as.character(fit$call$formula)[3]
        w <- unique(z$msg$warnings)
        z <- "confint(fit)"
        z <- evals(z, env = environment(), graph.dir = "/tmp/R_pander")[[1]]
        ci <- z$result
        w <- c(w, unique(z$msg$warnings))
        w <- paste(w, collapse = " / ")
        b <- cbind(coef(fit), ci)
        if (fam == "binomial") {
          b <- exp(b)
          colnames(b)[1] <- "odds ratio"
        } else {
          colnames(b)[1] <- "beta"
        }
        pv <- coef(summary(fit))[, 4]
        s <- ifelse(pv > 0.1, NA, ifelse(pv > 0.05, "≤0.1", ifelse(pv > 0.01,
               "≤0.05", ifelse(pv > 0.001, "≤0.01" ,"≤0.001"))))
        s2 <- ifelse(grepl("Groupe", rownames(b)) & pv <= 0.1, 1, NA)
        s[1] <- NA
        b <- cbind(b, `p-value` = pv, significativity = s, signif_group = s2)
        na <- rep(NA, nrow(b) - 1)
        r <- data.frame(
          `dependent variable` = c(y, na),
          `independent variables` = c(x_fml, na),
          `number of observations` = c(nrow(fit$model), na),
          coefficient = rownames(b)
        )
        cbind(r, b, `fit warnings` = c(w, na))
      }))
    })
  })
})
for (k in 1:4) {
  for (z in names(mvreg[[k]])) {
    if (k %in% 1:2) {
      s <- paste0(k + 1, "_reg")
    } else if (k == 3) {
      s <- "complet"
    } else {
      s <- "stepwise"
    }
    f <- file.path(outdir, paste0("regressions_multivar_", s, "_", z, ".xlsx"))
    write_xlsx(mvreg[[k]][[z]], f)
  }
}
rm(f, k, z)

# Session Info
sink(file.path(outdir, "sessionInfo.txt"))
print(sessionInfo(), locale = FALSE)
sink()
