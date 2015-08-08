namespace FinAnParser
{
    class Calculator
    {
        public double ResultatKoeficientaFinNezavisimosti = 0.0;
        public double ResultatRentabelnostiProdag = 0.0;
        public double ResultatRentabelnostiDeyatelnosti = 0;

        public double ItogIzmeneniyaViruchki = 0;
        public double ItogKoeficientaTekucsheyLikvidnosti = 0;
        public double ItogKoeficientaFinNezavisimosti = 0;
        public double ItogChistihActivov = 0;
        public double ItogRentabelnostiProdag = 0;
        public double ItogRentabelnostiDeyatelnosti = 0;

        public double VzvesheniyItogIzmeneniyaViruchki = 0;
        public double VzvesheniyItogKoeficientaTekucsheyLikvidnosti = 0;
        public double VzvesheniyItogKoeficientaFinNezavisimosti = 0;
        public double VzvesheniyItogChistihActivov = 0;
        public double VzvesheniyItogRentabelnostiProdag = 0;
        public double VzvesheniyItogRentabelnostiDeyatelnosti = 0;

        public void PodschetIzmeneniyaViruchki(double izmenenieViruchki)
        {
            // Свыше 10
            if (izmenenieViruchki > 10)
            {
                ItogIzmeneniyaViruchki = 1;
            }

            // От 0 до 10
            if ((izmenenieViruchki > 0) && (izmenenieViruchki <= 10))
            {
                ItogIzmeneniyaViruchki = 0.8;
            }

            // От -25 до 0
            if ((izmenenieViruchki > -25) && (izmenenieViruchki <= 0))
            {
                ItogIzmeneniyaViruchki = 0.6;
            }

            // От 0.5 до 0.8
            if (izmenenieViruchki < -25)
            {
                ItogIzmeneniyaViruchki = 0.4;
            }

            VzvesheniyItogIzmeneniyaViruchki = ItogIzmeneniyaViruchki * 0.2;
        }

        public void PodschetKoeficientaTekucsheyLikvidnosti(double koeficientTekucsheyLikvidnosti)
        {
            // Свыше 1.2
            if (koeficientTekucsheyLikvidnosti > 1.2)
            {
                ItogKoeficientaTekucsheyLikvidnosti = 1;
            }

            // От 1 до 1.2
            if ((koeficientTekucsheyLikvidnosti > 1) && (koeficientTekucsheyLikvidnosti <= 1.2))
            {
                ItogKoeficientaTekucsheyLikvidnosti = 0.8;
            }

            // От 0.8 до 1
            if ((koeficientTekucsheyLikvidnosti > 0.8) && (koeficientTekucsheyLikvidnosti <= 1))
            {
                ItogKoeficientaTekucsheyLikvidnosti = 0.6;
            }

            // От 0.5 до 0.8
            if ((koeficientTekucsheyLikvidnosti > 0.5) && (koeficientTekucsheyLikvidnosti <= 0.8))
            {
                ItogKoeficientaTekucsheyLikvidnosti = 0.4;
            }

            // Менее 0.5
            if (koeficientTekucsheyLikvidnosti < 0.5)
            {
                ItogKoeficientaTekucsheyLikvidnosti = 0.2;
            }

            VzvesheniyItogKoeficientaTekucsheyLikvidnosti = ItogKoeficientaTekucsheyLikvidnosti * 0.2;

        }

        public void PodschetKoeficientaFinNezavisimosti(double sobstvenniyKapital, double valyutniyBalans, string okvedVidDeyatelnosti)
        {
            ResultatKoeficientaFinNezavisimosti = sobstvenniyKapital / valyutniyBalans;

            if (okvedVidDeyatelnosti == "Производство")
            {
                // От 0,5
                if (ResultatKoeficientaFinNezavisimosti > 0.5)
                {
                    ItogKoeficientaFinNezavisimosti = 1;
                }

                // От 0.4 до 0.5
                if ((ResultatKoeficientaFinNezavisimosti > 0.4) && (ResultatKoeficientaFinNezavisimosti <= 0.5))
                {
                    ItogKoeficientaFinNezavisimosti = 0.8;
                }

                // От 0 до 0.4
                if ((ResultatKoeficientaFinNezavisimosti > 0) && (ResultatKoeficientaFinNezavisimosti <= 9.0))
                {
                    ItogKoeficientaFinNezavisimosti = 0.6;
                }

                // Менее 0
                if (ResultatKoeficientaFinNezavisimosti < 0)
                {
                    ItogKoeficientaFinNezavisimosti = 0.4;
                }

            }

            if (okvedVidDeyatelnosti == "Торговля")
            {
                // От 0,3
                if (ResultatKoeficientaFinNezavisimosti > 0.3)
                {
                    ItogKoeficientaFinNezavisimosti = 1;
                }

                // От 0.4 до 0.5
                if ((ResultatKoeficientaFinNezavisimosti > 0.15) && (ResultatKoeficientaFinNezavisimosti <= 0.3))
                {
                    ItogKoeficientaFinNezavisimosti = 0.8;
                }

                // От 0 до 0.4
                if ((ResultatKoeficientaFinNezavisimosti > 0) && (ResultatKoeficientaFinNezavisimosti <= 0.15))
                {
                    ItogKoeficientaFinNezavisimosti = 0.6;
                }

                // Менее 0
                if (ResultatKoeficientaFinNezavisimosti < 0)
                {
                    ItogKoeficientaFinNezavisimosti = 0.4;
                }
            }

            VzvesheniyItogKoeficientaFinNezavisimosti = ItogKoeficientaFinNezavisimosti * 0.2;

        }

        public void PodschetChistihActivov(double chistieActiviNaNachaloPerioda, double chistieActiviNaKonecPerioda, double chistieActiviIzmeneniya)
        {
            // Увеличение (ЧА-положительное значение)
            if ((chistieActiviNaNachaloPerioda < chistieActiviNaKonecPerioda) & chistieActiviIzmeneniya > 0)
            {
                ItogChistihActivov = 1;
            }

            // Уменьшение (ЧА-положительное значение)
            if ((chistieActiviNaNachaloPerioda > chistieActiviNaKonecPerioda) & chistieActiviIzmeneniya > 0)
            {
                ItogChistihActivov = 0.8;
            }

            // Увеличение (ЧА-отрицательное значение)
            if ((chistieActiviNaNachaloPerioda < chistieActiviNaKonecPerioda) & chistieActiviIzmeneniya < 0)
            {
                ItogChistihActivov = 0.6;
            }

            // Увеличение (ЧА-отрицательное значение)
            if ((chistieActiviNaNachaloPerioda > chistieActiviNaKonecPerioda) & chistieActiviIzmeneniya < 0)
            {
                ItogChistihActivov = 0.4;
            }

            VzvesheniyItogChistihActivov = ItogChistihActivov * 0.2;
        }

        public void PodschetRentabelnostiProdag(double pribilUbitokOtProdag, double viruchka)
        {
            if (viruchka != 0)
            {
                ResultatRentabelnostiProdag = (pribilUbitokOtProdag / viruchka) * 100;
            }

            // От 18%
            if (ResultatRentabelnostiProdag > 18.0)
            {
                ItogRentabelnostiProdag = 1;
            }

            // От 9% до 18%
            if ((ResultatRentabelnostiProdag > 9.0) && (ItogRentabelnostiProdag <= 18.0))
            {
                ItogRentabelnostiProdag = 0.8;
            }

            // От 0% до 9%
            if ((ResultatRentabelnostiProdag > 0) && (ItogRentabelnostiProdag <= 9.0))
            {
                ItogRentabelnostiProdag = 0.6;
            }

            // Менее 0%
            if (ResultatRentabelnostiProdag < 0)
            {
                ItogRentabelnostiProdag = 0;
            }

            VzvesheniyItogRentabelnostiProdag = ItogRentabelnostiProdag * 0.05;

        }

        public void PodschetRentabelnostiDeyatelnosti(double viruchkaPlusMinus, double viruchka)
        {
            if (viruchka != 0)
            {
                ResultatRentabelnostiDeyatelnosti = (viruchkaPlusMinus / viruchka) * 100;
            }

            // От 12%
            if (ResultatRentabelnostiDeyatelnosti > 12.0)
            {
                ItogRentabelnostiDeyatelnosti = 1;
            }

            // От 6% до 12%
            if ((ResultatRentabelnostiDeyatelnosti > 6.0) && (ResultatRentabelnostiDeyatelnosti <= 12.0))
            {
                ItogRentabelnostiDeyatelnosti = 0.8;
            }

            // От 0% до 6%
            if ((ResultatRentabelnostiDeyatelnosti > 0) && (ResultatRentabelnostiDeyatelnosti <= 6.0))
            {
                ItogRentabelnostiDeyatelnosti = 0.6;
            }

            // Менее 0%
            if (ResultatRentabelnostiDeyatelnosti < 0)
            {
                ItogRentabelnostiDeyatelnosti = 0;
            }

            VzvesheniyItogRentabelnostiDeyatelnosti = ItogRentabelnostiDeyatelnosti * 0.15;
        }
    }
}