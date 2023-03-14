library(data.table)
library(readxl)


# their configuration
data_biskuvi <-
  read_excel("C:/Users/yagiz.yaman/Desktop/RealDataBiskuvi.xlsx",
             sheet = "OrderData")

data_biskuvi <- as.data.table(data_biskuvi)

setnames(data_biskuvi, "Kaynak Adres", "cell")
setnames(data_biskuvi, "Malzeme", "sku")
setnames(data_biskuvi, "Emir No", "emir_no")

data_biskuvi <- data_biskuvi[, .(emir_no, sku, cell)]
data_biskuvi <- data_biskuvi[order(emir_no, cell)]

data_biskuvi[, shifted_cell := shift(cell), emir_no]



## distance_matrix

distance_matrix <-
  read_excel(
    "C:/Users/yagiz.yaman/Desktop/distancematrix.xlsx",
    sheet = "XXX"
  )
distance_matrix <- as.data.table(distance_matrix)
#names(distance_matrix)
setnames(distance_matrix, "...1", "empty")#cellVeSkucell
distance_matrix <- distance_matrix[empty %like% "SH"]


### create data.table from matrix distance_matrix
holder <- distance_matrix[, empty]
my_holder <- rep(holder, nrow(distance_matrix))

cell_holder <- colnames(distance_matrix)
cell_holder <- cell_holder[-1]
cells <- rep(cell_holder, nrow(distance_matrix))

my_holder <- my_holder[order(my_holder)]


my_dt <- data.table(cell1 = my_holder, cell2 = cells)

holder <- holder[order(holder)]

x <- c()
for (i in holder){
  x <- c(x, 
         as.numeric(distance_matrix[empty == i, 
                                    -c("empty")]))
}

my_dt[, distances := x]


my_dt[, cell1 := gsub(" ", "", cell1)]
my_dt[, cell2 := gsub(" ", "", cell2)]

#merge
data_biskuvi[is.na(shifted_cell), shifted_cell:= cell]

setkey(data_biskuvi, "cell", "shifted_cell")
setkey(my_dt, "cell1", "cell2")



my_dt <- my_dt[data_biskuvi]
setcolorder(my_dt, c("emir_no", "sku", "cell1", "cell2", "distances"))

my_dt <- my_dt[order(emir_no, cell1)]

## distance_to_pool

distance_to_pool <- 
  read_excel(
    "C:/Users/yagiz.yaman/Desktop/dist_to_pool.xlsx",
    sheet = "DistancePool"
  )

distance_to_pool <- as.data.table(distance_to_pool)

distance_to_pool[, cell2 := Cell]

setnames(distance_to_pool, "Cell", "cell1")

setkey(distance_to_pool, "cell1", "cell2")
setkey(eti, "cell1", "cell2")

my_dt <- distance_to_pool[my_dt]

setcolorder(my_dt, c("emir_no", "sku", "cell1", "cell2", "distances",
                          "Distance_to_Pool_Area"))

my_dt <- my_dt[order(emir_no, cell1)]

my_dt[distances==0, distances := Distance_to_Pool_Area]


# shop_floor configuration

sol_group <-   read_excel("C:/Users/yagiz.yaman/Desktop/Solutiongroup.xlsx")
sol_group <- as.data.table(sol_group)
setnames(sol_group, "ALAN", "cell")
setnames(sol_group, "SKU", "sku")


data_biskuvi <-
  read_excel("C:/Users/yagiz.yaman/Desktop/RealDataBiskuvi.xlsx",
             sheet = "OrderData")

data_biskuvi <- as.data.table(data_biskuvi)

setnames(data_biskuvi, "Kaynak Adres", "cell")
setnames(data_biskuvi, "Malzeme", "sku")
setnames(data_biskuvi, "Emir No", "emir_no")


data_biskuvi <- data_biskuvi[, .(emir_no, sku)]

# merge two

data_biskuvi[, sku := as.character(sku)]

sol_group[, cell := gsub(" ", "", cell)]
sol_group[, sku := gsub(" ", "", sku)]


setkey(sol_group, "sku")
setkey(data_biskuvi, "sku")

shop_floor <- sol_group[data_biskuvi]
setcolorder(shop_floor, c("emir_no", "sku", "cell"))

shop_floor <- shop_floor[order(emir_no, cell)]

shop_floor[, shifted_cell := shift(cell), emir_no]

shop_floor[is.na(shifted_cell), shifted_cell := cell]


# calculate distance

## distance_matrix

distance_matrix <-
  read_excel(
    "C:/Users/yagiz.yaman/Desktop/distancematrix.xlsx",
    sheet = "XXX"
  )
distance_matrix <- as.data.table(distance_matrix)
#names(distance_matrix)
setnames(distance_matrix, "...1", "empty")#cellVeSkucell
distance_matrix <- distance_matrix[empty %like% "SH"]


### create data.table from matrix distance_matrix
holder <- distance_matrix[, empty]
my_holder <- rep(holder, nrow(distance_matrix))

cell_holder <- colnames(distance_matrix)
cell_holder <- cell_holder[-1]
cells <- rep(cell_holder, nrow(distance_matrix))

my_holder <- my_holder[order(my_holder)]


my_new <- data.table(cell1 = my_holder, cell2 = cells)

holder <- holder[order(holder)]

x <- c()
for (i in holder){
  x <- c(x, 
         as.numeric(distance_matrix[empty == i, 
                                    -c("empty")]))
}

my_new[, distances := x]


my_new[, cell1 := gsub(" ", "", cell1)]
my_new[, cell2 := gsub(" ", "", cell2)]


#merge

setkey(shop_floor, "cell", "shifted_cell")
setkey(my_new, "cell1", "cell2")

shop_floor <- my_new[shop_floor]
setcolorder(shop_floor, c("emir_no", "sku", "cell1", "cell2", "distances"))

shop_floor <- shop_floor[order(emir_no, cell1)]


# distance pool

## distance_to_pool

distance_to_pool <- 
  read_excel(
    "C:/Users/yagiz.yaman/Desktop/dist_to_pool.xlsx",
    sheet = "DistancePool"
  )

distance_to_pool <- as.data.table(distance_to_pool)

distance_to_pool[, cell2 := Cell]

setnames(distance_to_pool, "Cell", "cell1")

setkey(distance_to_pool, "cell1", "cell2")
setkey(shop_floor, "cell1", "cell2")

shop_floor <- distance_to_pool[shop_floor]

setcolorder(shop_floor, c("emir_no", "sku", "cell1", "cell2", "distances",
                        "Distance_to_Pool_Area"))

shop_floor <- shop_floor[order(emir_no, cell1)]

shop_floor[distances==0, distances := Distance_to_Pool_Area]



print(paste0("The improvement is ", 
      round(100*
        (my_dt[, sum(distances, na.rm=TRUE)]-shop_floor[, sum(distances, na.rm=TRUE)])/
        my_dt[, sum(distances, na.rm=TRUE)],2), 
      "%. The distance taken in the existing configuration is ", 
      eti[, sum(distances, na.rm=TRUE)], 
      " meters. The distance taken in the improved configuration is ", 
      shop_floor[, sum(distances, na.rm=TRUE)], " meters."))

