import React from 'react';
import Tabs from './tabs';
import DisplayImage from './displayImage';
import '../assets/css/productPage.css';
import GentleIcon from '../assets/images/displayImages/Compound Path_1.png';
import SafetyIcon from '../assets/images/displayImages/Group_1.png';
import GentleRating from '../assets/images/displayImages/Text_1.png';
import SafetyRating from '../assets/images/displayImages/Text_2.png';

const ProductPage = () => {

    const page = {
        position:'fixed',
        bottom: '0px'
    }

    return (
        <div>
            <DisplayImage/>
            <div className="product-page-gentle-safety-rating">
                <div className="product-page-gentle-rating">GENTLE RATING
                    <span><img src={GentleIcon}/><img src={GentleRating}/></span>
                </div>
                <div className="product-page-safety-rating">SAFETY RATING
                    <span><img src={SafetyIcon}/><img src={SafetyRating}/></span>
                </div>
            </div>
            <div style={page}>
                <Tabs/>
            </div>
        </div>
    )
};

export default ProductPage;