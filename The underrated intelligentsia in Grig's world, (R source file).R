library(readxl)
library(ggplot2)
library(gridExtra)
library(grid)
library(ggplotify)

#"Ծերուկ"
raw_data_1 <- read_excel("Excel.xlsx", sheet = "Ծերուկ")
clean_data_1 <- na.omit(raw_data_1[, c("Negative", "Positive")])

table_data_1 <- clean_data_1
colnames(table_data_1) <- c("Բացասական", "Դրական")

custom_theme_1 <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_1 <- tableGrob(
  table_data_1,
  rows = NULL,
  theme = custom_theme_1
)

word_table_1$widths <- unit(c(0.4, 0.6), "npc")

comparison_1 <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(clean_data_1$Negative)) * 2,
    sum(!is.na(clean_data_1$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_1 <- ggplot(comparison_1, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) + 
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial" 
  ) +
  scale_fill_identity() +
  labs(
    title = "Ծերուկ",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_1$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3))) 

spacer_1 <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_1 <- grid.arrange(
  weighted_plot_1,
  spacer_1,
  word_table_1,
  nrow = 3,
  heights = c(3, 0.6, 3)
)






#"Սոֆա"
raw_data_2 <- read_excel("Excel.xlsx", sheet = "Սոֆա")
clean_data_2 <- na.omit(raw_data_2[, c("Negative", "Positive")])

table_data_2 <- clean_data_2
colnames(table_data_2) <- c("Բացասական", "Դրական")

custom_theme_2 <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_2 <- tableGrob(
  table_data_2,
  rows = NULL,
  theme = custom_theme_2
)

word_table_2$widths <- unit(c(0.4, 0.6), "npc")

comparison_2 <- data.frame(
  Category = c("Դրական", "Բացասական"),
  Score = c(
    sum(!is.na(clean_data_2$Negative)) * 2,
    sum(!is.na(clean_data_2$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_2 <- ggplot(comparison_2, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +  
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial" 
  ) +
  scale_fill_identity() +
  labs(
    title = "Սոֆա",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_2$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3))) 

spacer_2 <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))  

final_layout_2 <- grid.arrange(
  weighted_plot_2,
  spacer_2,
  word_table_2,
  nrow = 3,
  heights = c(3, 0.6, 3)
)



#"Սամվել"
raw_data_samvel <- read_excel("Excel.xlsx", sheet = "Սամվել")
clean_data_samvel <- na.omit(raw_data_samvel[, c("Negative", "Positive")])

table_data_samvel <- clean_data_samvel
colnames(table_data_samvel) <- c("Բացասական", "Դրական")

custom_theme_samvel <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_samvel <- tableGrob(
  table_data_samvel,
  rows = NULL,
  theme = custom_theme_samvel
)

word_table_samvel$widths <- unit(c(0.4, 0.6), "npc")

comparison_samvel <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(clean_data_samvel$Negative)) * 2,
    sum(!is.na(clean_data_samvel$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_samvel <- ggplot(comparison_samvel, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial"
  ) +
  scale_fill_identity() +
  labs(
    title = "Սամվել",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_samvel$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3)))

spacer_samvel <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_samvel <- grid.arrange(
  weighted_plot_samvel,
  spacer_samvel,
  word_table_samvel,
  nrow = 3,
  heights = c(3, 0.6, 3)
)




#"Երանյան Աշոտ"
raw_data_ashot <- read_excel("Excel.xlsx", sheet = "Երանյան Աշոտ")
clean_data_ashot <- na.omit(raw_data_ashot[, c("Negative", "Positive")])

table_data_ashot <- clean_data_ashot
colnames(table_data_ashot) <- c("Բացասական", "Դրական") 

custom_theme_ashot <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_ashot <- tableGrob(
  table_data_ashot,
  rows = NULL,
  theme = custom_theme_ashot
)

word_table_ashot$widths <- unit(c(0.4, 0.6), "npc") 

comparison_ashot <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(clean_data_ashot$Negative)) * 2,
    sum(!is.na(clean_data_ashot$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_ashot <- ggplot(comparison_ashot, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial"
  ) +
  scale_fill_identity() +
  labs(
    title = "Երանյան Աշոտ",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_ashot$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3)))

spacer_ashot <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_ashot <- grid.arrange(
  weighted_plot_ashot,
  spacer_ashot,
  word_table_ashot,
  nrow = 3,
  heights = c(3, 0.6, 3)
)




#"Կառոբկա Աշոտ"
raw_data_karobka <- read_excel("Excel.xlsx", sheet = "Կառոբկա Աշոտ")
clean_data_karobka <- na.omit(raw_data_karobka[, c("Negative", "Positive")])

n_missing_rows <- max(0, 4 - nrow(clean_data_karobka))
if (n_missing_rows > 0) {
  empty_rows <- data.frame(Negative = rep("", n_missing_rows), Positive = rep("", n_missing_rows))
  clean_data_karobka <- rbind(clean_data_karobka, empty_rows)
}

table_data_karobka <- clean_data_karobka
colnames(table_data_karobka) <- c("Բացասական", "Դրական")

custom_theme_karobka <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_karobka <- tableGrob(
  table_data_karobka,
  rows = NULL,
  theme = custom_theme_karobka
)

word_table_karobka$widths <- unit(c(0.4, 0.6), "npc")

comparison_karobka <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(raw_data_karobka$Negative)) * 2,
    sum(!is.na(raw_data_karobka$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_karobka <- ggplot(comparison_karobka, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial"
  ) +
  scale_fill_identity() +
  labs(
    title = "Կառոբկա Աշոտ",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_karobka$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3)))

spacer_karobka <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_karobka <- grid.arrange(
  weighted_plot_karobka,
  spacer_karobka,
  word_table_karobka,
  nrow = 3,
  heights = c(3, 0.6, 3)
)




#"Օպոզիցիա Տիգրան"
raw_data_opozicia <- read_excel("Excel.xlsx", sheet = "Օպոզիցիա Տիգրան")
clean_data_opozicia <- na.omit(raw_data_opozicia[, c("Negative", "Positive")])

n_missing_rows <- max(0, 3 - nrow(clean_data_opozicia))
if (n_missing_rows > 0) {
  empty_rows <- data.frame(Negative = rep("", n_missing_rows), Positive = rep("", n_missing_rows))
  clean_data_opozicia <- rbind(clean_data_opozicia, empty_rows)
}

table_data_opozicia <- clean_data_opozicia
colnames(table_data_opozicia) <- c("Բացասական", "Դրական")

custom_theme_opozicia <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "dodgerblue2"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_opozicia <- tableGrob(
  table_data_opozicia,
  rows = NULL,
  theme = custom_theme_opozicia
)

word_table_opozicia$widths <- unit(c(0.4, 0.6), "npc")

comparison_opozicia <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(raw_data_opozicia$Negative)) * 2,
    sum(!is.na(raw_data_opozicia$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_opozicia <- ggplot(comparison_opozicia, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial"
  ) +
  scale_fill_identity() +
  labs(
    title = "Օպոզիցիա Տիգրան",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_opozicia$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3)))

spacer_opozicia <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_opozicia <- grid.arrange(
  weighted_plot_opozicia,
  spacer_opozicia,
  word_table_opozicia,
  nrow = 3,
  heights = c(3, 0.6, 3)
)




#"Գրիգի հայր"
raw_data_grigi <- read_excel("Excel.xlsx", sheet = "Գրիգի հայր")
clean_data_grigi <- na.omit(raw_data_grigi[, c("Negative", "Positive")])

n_missing_rows <- max(0, 3 - nrow(clean_data_grigi))
if (n_missing_rows > 0) {
  empty_rows <- data.frame(Negative = rep("", n_missing_rows), Positive = rep("", n_missing_rows))
  clean_data_grigi <- rbind(clean_data_grigi, empty_rows)
}

table_data_grigi <- clean_data_grigi
colnames(table_data_grigi) <- c("Բացասական", "Դրական")

custom_theme_grigi <- ttheme_minimal(
  core = list(
    bg_params = list(fill = c("#F8F8F8", "white"), col = NA),
    fg_params = list(
      hjust = 0,
      x = 0.05,
      fontsize = 11,
      fontfamily = "Arial",
      lineheight = 1.1
    )
  ),
  colhead = list(
    bg_params = list(fill = "deepskyblue"),
    fg_params = list(
      col = "white",
      fontface = "bold",
      fontsize = 12,
      fontfamily = "Arial"
    )
  ),
  padding = unit(c(6, 10), "mm")
)

word_table_grigi <- tableGrob(
  table_data_grigi,
  rows = NULL,
  theme = custom_theme_grigi
)

word_table_grigi$widths <- unit(c(0.4, 0.6), "npc")

comparison_grigi <- data.frame(
  Category = c("Բացասական", "Դրական"),
  Score = c(
    sum(!is.na(raw_data_grigi$Negative)) * 2,
    sum(!is.na(raw_data_grigi$Positive)) * 1
  ),
  Color = c("dodgerblue4", "deepskyblue")
)

weighted_plot_grigi <- ggplot(comparison_grigi, aes(x = Category, y = Score, fill = Color)) +
  geom_col(width = 0.5) +
  geom_text(
    aes(label = Score),
    vjust = -0.8,
    size = 6,
    fontface = "bold",
    family = "Arial"
  ) +
  scale_fill_identity() +
  labs(
    title = "Գրիգի հայր",
    y = "Գնահատական",
    x = NULL
  ) +
  ylim(0, max(comparison_grigi$Score) * 1.4) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, hjust = 0.5, family = "Arial", margin = margin(b = 20)),
    axis.text = element_text(size = 12, family = "Arial"),
    axis.title.y = element_text(size = 14, family = "Arial", margin = margin(r = 15)),
    panel.grid.major.x = element_blank(),
    plot.margin = margin(t = 20, r = 30, b = 10, l = 30)
  ) +
  scale_x_discrete(expand = expansion(mult = c(0.3, 0.3)))

spacer_grigi <- rectGrob(gp = gpar(col = NA, fill = NA), height = unit(2, "cm"))

final_layout_grigi <- grid.arrange(
  weighted_plot_grigi,
  spacer_grigi,
  word_table_grigi,
  nrow = 3,
  heights = c(3, 0.6, 3)
)


