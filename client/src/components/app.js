import React from 'react';
import '../assets/css/footer_landing_sign_in.css';
import LandingPage from './landing_page';
import SignIn from './sign_in';
import ProductFinder from './product_finder';
import { Route } from 'react-router-dom';
import MyAccountPage from './myAccount';
import ExpandedMenuWelcome from './expandedMenuWelcome';
import LookUpIngredient from './lookUpIngredient';
import CheckProducts from './checkProducts';
import ProductPage from './productPage';
import SafetyRating from './safetyRating';
import GentleRating from './gentleRating';
import DisplayAllStandAlone from './displayAll_standAlone.js';
import CreateAccount from './create_account';
import OurTeam from './ourTeam';
import FAQ from './FAQ';
import AboutTheSite from './aboutTheSite';
import ExpandedMenu from './expandedMenuWelcome';

const App = () => {
    return (
        <section>
            <Route exact path="/" component={LandingPage} />
            <Route path="/sign_in" component={SignIn} />
            <Route path="/product_finder" component={ProductFinder} />
            <Route path="/myAccount" component={MyAccountPage} />
            <Route path="/search_ingredient" component={LookUpIngredient} />
            <Route path="/check_products" component={CheckProducts} />
            <Route path="/product/:id?" component={ProductPage} />
            <Route path="/safety_rating" component={SafetyRating} />
            <Route path="/gentle_rating" component={GentleRating} />
            <Route path="/gentle_products" component={DisplayAllStandAlone} />
            <Route path="/create_account" component={CreateAccount} />
            <Route path="/our_team" component={OurTeam} />
            <Route path="/faq" component={FAQ} />
            <Route path="/about_the_site" component={AboutTheSite} />
            <Route path="/expanded_menu" component={ExpandedMenu} />
        </section>
    )
};

export default App;

//componentwill mount, scroll  up to the top of the page