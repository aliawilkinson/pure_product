import React, { Component } from 'react';
import { Link } from 'react-router-dom';
import Footer from './footer/footer';
import DisplayAllProducts from './displayProduct/displayAllProducts';
import gentle_rating from '../assets/images/landing_page_icons/gentle_feather.png';
import foundation from '../assets/images/landing_page_icons/foundation.png';
import lipstick from '../assets/images/landing_page_icons/icons/cross2.png';
import safety_icon from '../assets/images/landing_page_icons/icons/cross.png';
import product from '../assets/images/landing_page_icons/icons/gentle2.png';
import Header from './header/header';
import '../assets/css/landing_page.css';
import axios from 'axios';

class LandingPage extends Component {
    constructor(props) {
        super(props);
        this.state = {
            data: {
                data: null
            }
        }
    }

    async componentDidMount() {

        var id = this.props.match.params.id;
        await axios.post(`/server/api_get_product_by_categories.php`)// can add another parameter to send to backend for queries, object

            .then(res => {
                this.setState({
                    data: res.data
                })
            })
    }

    render() {
        const { data } = this.state;
        return (
            <section className="landing_page">
                <Header history={this.props.history} />
                <div className="landing-page-button-container">
                    <div className="landing-two-container">
                        <Link to="/browse_products">
                            <div className="landing_button products">
                                <div className="wrap-landing-img">
                                    <img className="landing_img product" src={product} />
                                    <span className="top-span">Browse</span>
                                    <span className="top-span">Products</span>
                                </div>
                            </div>
                        </Link>
                        <Link to="/product_analyzer">
                            <div className="landing_button safety">
                                <div className="wrap-landing-img">
                                    <img className="landing_img lipstick" src={lipstick} />
                                    <span className="top-span">Analyze</span>
                                    <span className="top-span">Ingredients</span>
                                </div>
                            </div>
                        </Link>
                    </div>
                    <div className="landing-two-container">
                        <Link to="/safety_rating">
                            <div className="landing_button safety_rating">
                                <div className="wrap-landing-img">
                                    <img className="landing_img safety_icon" src={safety_icon} />
                                    <span className="safety-span">Safety</span>
                                    <span className="safety-span">Rating</span>
                                </div>
                            </div>
                        </Link>
                        <Link to="/gentle_rating">
                            <div className="landing_button gentle_rating">
                                <div className="wrap-landing-img">
                                    <img className="landing_img gentle" src={gentle_rating} />
                                    <span className="gentle-span">Gentle</span>
                                    <span className="gentle-span">Rating</span>
                                </div>
                            </div>
                        </Link>
                    </div>
                </div>
                <Link to="/product_wizard">
                    <div className="product_finder_landing_button waves-effect waves-light">
                        <span>Foundation Finder</span>
                    </div>
                </Link>
                <DisplayAllProducts data={data} />
                <Footer />
            </section>
        )
    }
}

export default LandingPage;