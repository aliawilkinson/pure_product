import React from 'react';
import { Link } from 'react-router-dom';
import '../../assets/css/header.css';
import MyAccountPicture from '../../assets/images/menu_icons/sign_in.png';
import CreateAccountPicture from '../../assets/images/menu_icons/sign_up.png';
import CheckProductSafety from '../../assets/images/landing_page_icons/icons/cross2.png';
import LookIngredients from '../../assets/images/landing_page_icons/icons/gentle2.png';
import MethodologyPicutre from '../../assets/images/menu_icons/home.png';
import DiscoverProducts from '../../assets/images/menu_icons/search.png';
import leaf from '../../assets/images/landing_page_icons/icons/vegan.png';
import home from '../../assets/images/home.png';
import '../../assets/css/expandedMenuStyle.css';

function ExpandedMenuWelcome(props) {

    return (
        <div className={`expanded-menu ${props.className}`}>
            <div className="menu-container">
                {/* <Link to="/myAccount" >
                    <div className="menu-item-cont">
                        <img className="menu-item-my-account" src={MyAccountPicture} align="middle"/><p className="exp-menu-titles">My Account</p>
                    </div>
                </Link> */}

                {/* <Link to="/create_account" >
                    <div className="menu-item-cont">
                        <img className="menu-item-create-account" src={CreateAccountPicture} align="middle"/><p className="exp-menu-titles">Create Account</p>
                    </div>
                </Link> */}

                <Link to="/product_analyzer" >
                    <div className="menu-item-cont">
                        <img className="menu-item-check-products" src={CheckProductSafety} align="middle" /><p className="exp-menu-titles">Analyze Product Ingredients</p>
                    </div>
                </Link>

                <Link to="/browse_products">
                    <div className="menu-item-cont">
                        <img className="menu-item-single-ingredient-search" src={LookIngredients} width="10%" align="middle" /><p className="exp-menu-titles">Browse Products</p>
                    </div>
                </Link>
                <Link to="/product_wizard">
                    <div className="menu-item-cont">
                        <img src={leaf} /><p className="product-finder-exp-menu-title exp-menu-titles">Personal Product Finder</p>
                    </div>
                </Link>
                <Link to="/" onClick={location.reload}>
                    <div className="menu-item-cont">
                        <img className="home" src={home} /><p className="product-finder-exp-menu-title exp-menu-titles">Home</p>
                    </div>
                </Link>

                {/* <Link to="/check_products">
                    <div className="menu-item-cont">
                        <img className="menu-item-discover-products" src={DiscoverProducts} align="middle" />Discover Products</p>
                    </div>
                </Link> */}
            </div>
        </div>
    )
};

export default ExpandedMenuWelcome;